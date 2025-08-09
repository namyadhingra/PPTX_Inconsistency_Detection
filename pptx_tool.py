# Activate the virtual environment "venv" using "venv\Scripts\activate" and install the dependencies.
# Dependencies in the file "dependencies.txt", run 
# pip install -r dependencies.txt 
# in the terminal.

# Section 1: Imports and Setup

import os, re, json, argparse, tempfile
from tqdm import tqdm
from pptx import Presentation
from PIL import Image
import pytesseract
import dateparser
import google.generativeai as genai
from pathlib import Path

GEMINI_MODEL = "gemini-2.5-flash"
genai.configure(api_key=os.environ.get("my_api_key"))
model = genai.GenerativeModel(GEMINI_MODEL)

# Section 2: Helper functions

# 1. PPTX helper function (for text-based)
def extract_text_from_pptx(pptx_path):
    """This is a helper function that extracts text and any embedded images, from a .pptx file"""
    prs = Presentation(pptx_path)
    # prs is the Presentation object that represents the .pptx file
    slides = []
    for i, slide in enumerate(prs.slides, start=1):
        text_chunks=[]
        images=[]
        for shape in slide.shapes:
            if hasattr(shape,"text") and shape.text.strip():
                # "hasattr" means that the shape has a "text" attribute, and "shape.text.strip()" means that the text is not empty
                # "strip()" is used to remove any leading or trailing whitespace from the text
                text_chunks.append(shape.text.strip())
            if hasattr(shape, "image"):
                try:
                    images.append(shape.image.blob)
                    # "shape.image.blob" is used to get the image bytes from the shape
                except Exception:
                    pass 
        slides.append({"slide":i , "text": " ".join(text_chunks) , "images":images})
        # "slide" is slide number, "text" is text extracted from the slide, "images" is a list of image bytes extracted from slide
    return slides

# 2. OCR helper function (for image-based)
def extract_text_from_images( img_folder):
    """This runs OCR on all images in the folder specified as "img_folder" """
    slides= []
    img_files = sorted( [f for f in os.listdir(img_folder) if f.lower().endswith((".png", ".jpg", ".jpeg"))])
    # Inclusion of .png, .jpg, when the assignment only has .jpeg files, makes the tool more generalized and scalable
    for index, img_file in enumerate(tqdm(img_files, desc="OCR images"), start=1):
        img_path = os.path.join(img_folder, img_file)
        text = pytesseract.image_to_string(Image.open(img_path)) #This is actual the OCR
        slides.append({"slide":index , "text":text , "images":[]})
    return slides

# 3. OCR for image bytes (image bytes = raw bytes)
def ocr_image_bytes (image_bytes):
    """This function takes image bytes and returns the OCR text"""
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        #This above line creates a temporary file with a .png suffix
        # delete=False means the file will not be deleted when closed, allowing us to read it
        tmp.write(image_bytes) #Does the actual writing of the image bytes to the temporary file
        tmp_path = tmp.name #tmp_path is set to name of temporary file
    try:
        text = pytesseract.image_to_string(Image.open(tmp_path))
        # "try" is for trying to open the temporary file and run OCR on it
    finally:
        os.remove(tmp_path) #deleting the temporary file after OCR is done
    return text

# 4. Normalization function for slides
def normalize_slides(slide):
    """This function normalizes the slide text by removing extra spaces and newlines, and extracts numbers/dates"""
    text = slide["text"] or ""
    if not text.strip() and slide["images"]:
        for img in slide["images"]:
            text += " " + ocr_image_bytes(img)
    num_regex = re.compile(r"(?:₹|\$|€)?\s*[-+]?\d{1,3}(?:[,\d{3}]+)?(?:\.\d+)?%?")
    # This above line is a regex that matches numbers with optional currency symbols and percentages
    # Also, it finds separators like commas and periods
    # Like the currency symbols help it to identify the currency, and the percentage symbol helps it to identify the percentage numbers
    numbers = list(set(num_regex.findall(text)))
    # This above line finds all the numbers in the text and removes duplicates
    # The set() function is used to remove duplicates, and then it is converted back to a list to maintain the order of numbers.
    dates=[]
    for word in text.split():
        dp = dateparser.parse(word, settings={'PREFER_DAY_OF_MONTH':'first'})
        # This line basically parses the word as a date, and if it is a valid date, it returns the date object
        # 'first' there means that if the date is ambiguous, it will prefer the first day of the month as default
        if dp: #If the date parser returns a valid date, it will be added to the list
            dates.append(dp.isoformat())
    return {"slide":slide["slide"] , "text": text.strip() , "numbers":numbers , "dates":list(set(dates))}


# Section 3: Rule-Based Detections


# 1. Detecting numeric conflicts
def detect_numeric_conflicts(slides):
    """This function finds same metrics with different values across slides."""
    metric_map = {} #Initializing an empty dictionary to store metrics and their values
    # The metric_map will have the metric as the key, and the value will be a dictionary with the value as the key and a list of slides as the value
    conflicts = [] #Initialized as an empty list

    for s in slides:
        pairs = re.findall(r"([A-Za-z %$₹€]{3,40})[:\-]?\s*([₹$€]?\s*\d[\d,\.]*%?)", s["text"])
        # Explanation of "([A-Za-z %$₹€]{3,40})[:\-]?\s*([₹$€]?\s*\d[\d,\.]*%?)"
        # ([A-Za-z %$₹€]{3,40}) matches a metric name that is 3 to 40 characters long, allowing letters, spaces, and currency symbols
        # [:\-]? matches an optional colon or hyphen after the metric name
        # \s* matches any whitespace after the colon or hyphen
        # ([₹$€]?\s*\d[\d,\.]*%?) matches a value that may start with a currency symbol (₹, $, €), followed by optional whitespace, a digit, and then 
        # any combination of digits, commas, periods, and an optional percentage sign
        # re.findall() returns a list of tuples, where each tuple contains the metric and its
        # corresponding value found in the text of the slide

        for metric, value in pairs:
            k = metric.strip().lower() #metric.strip().lower() removes leading/trailing spaces and converts to lowercase
            metric_map.setdefault(k, {}).setdefault(value, []).append(s["slide"]) 
            #Adds the slide to the list of slides for the given metric and value

    for metric, vals in metric_map.items():
        if len(vals) > 1:
            conflicts.append({"type": "numeric_conflict", "metric": metric, "values": vals})
    return conflicts

# 2. Detecting percentage sum issues
def detect_percent_sum_issues(slides, tolerance=2.0):
    """Checks if percentages in a slide sum to ~100%."""
    issues = []
    for s in slides:
        percents = re.findall(r"[-+]?\d{1,3}(?:\.\d+)?%", s["text"])
        #Explanation of "(r"[-+]?\d{1,3}(?:\.\d+)?%", s["text"])":
        # [-+]? matches an optional sign (either + or -)
        # \d{1,3} matches 1 to 3 digits
        # (?:\.\d+)? matches an optional decimal point followed by one or more digits
        # % matches the percentage sign
        # re.findall() returns a list of all matches of the regex in the text of the slide
        if len(percents) >= 2:
            total = sum(float(p.strip("%")) for p in percents)
            #Converts each percentage string to a float after stripping the '%' sign and sums them up
            if abs(total - 100.0) > tolerance:
                issues.append({"type": "percent_total_mismatch", "slide": s["slide"], "found_total": total, "details": percents})
                #if absolute diff between total and 100.0 is greater than the tolerance, it adds an issue
    return issues


# Section 4: Gemini Analysis

def extract_json_from_text(t):
    """Extract the first JSON array substring from text (crude but practical)."""
    if not isinstance(t, str):
        return None
    start = t.find('[')
    end = t.rfind(']')
    if start != -1 and end != -1 and end > start:
        return t[start:end+1]
    return None

def call_gemini(slide_objs):
    """Calls Gemini to detect contradictions and inconsistencies."""
    slide_texts = []
    for s in slide_objs:
        ppt_excerpt = (s.get("pptx_text") or "")[:1200]
        ocr_excerpt = (s.get("ocr_text") or "")[:1200]
        slide_texts.append(f"SLIDE {s['slide']}:\nPPTX: {ppt_excerpt}\nOCR: {ocr_excerpt}")

    prompt = (
        "Analyze the slides for factual or logical inconsistencies by comparing PPTX text and OCR text. "
        "Return a JSON array where each element has: issue_type, slides (list), excerpt, explanation. "
        "issue_type ∈ {numeric_conflict, contradictory_claim, timeline_mismatch, other}. "
        "If no issues, return [].\n\nSlides:\n" + "\n\n".join(slide_texts)
    )

    try:
        resp = model.generate_content(prompt)
        text = getattr(resp, "text", None) or str(resp)
    except Exception as e:
        return [{"type": "llm_error", "error": str(e)}]

    json_blob = extract_json_from_text(text)
    if not json_blob:
        # fallback: try parsing whole text
        try:
            return json.loads(text)
        except Exception:
            return [{"type": "parsing_error", "raw": text}]
    try:
        return json.loads(json_blob)
    except Exception as e:
        return [{"type": "parsing_error", "error": str(e), "raw": json_blob}]


def get_default_paths():
    base_dir = Path(__file__).resolve().parent  # script is in NOOGATASSIGNMENT/
    # the images are expected under NoogatAssignment/NoogatAssignment/
    images_dir = base_dir / "NoogatAssignment" / "NoogatAssignment"
    pptx_files = list(base_dir.rglob("*.pptx"))
    pptx_file = pptx_files[0] if pptx_files else None

    return str(pptx_file) if pptx_file else None, str(images_dir)


def main():
    default_pptx, default_images = get_default_paths()

    print("\n--- NOOGATASSIGNMENT Default Path Detection ---")
    print(f"Default PPTX:   {default_pptx if default_pptx else 'Not found'}")
    print(f"Default Images: {default_images if os.path.isdir(default_images) else 'Not found'}")
    print("------------------------------------------------")

    pptx_path = input(f"Enter PPTX path [{default_pptx}]: ").strip() or default_pptx
    images_path = input(f"Enter images folder [{default_images}]: ").strip() or default_images

    if not pptx_path or not os.path.isfile(pptx_path):
        print(f"❌ PPTX file not found: {pptx_path}")
        return
    if not images_path or not os.path.isdir(images_path):
        print(f"❌ Images folder not found: {images_path}")
        return

    print("\n✅ Paths confirmed:")
    print(f"PPTX:   {pptx_path}")
    print(f"Images: {images_path}")

    # Start processing: extract text from PPTX and images
    print("\n[1/5] Extracting text from PPTX slides...")
    raw_pptx_slides = extract_text_from_pptx(pptx_path)

    print("[2/5] Running OCR on slide images...")
    raw_image_slides = extract_text_from_images(images_path)

    # Normalize both sources
    print("[3/5] Normalizing slides (extracting numbers/dates)...")
    pptx_norm = [normalize_slides(s) for s in tqdm(raw_pptx_slides, desc="Normalizing PPTX")]
    ocr_norm = [normalize_slides(s) for s in tqdm(raw_image_slides, desc="Normalizing OCR")]

    # Merge per-slide (handle differing lengths)
    max_slides = max(len(pptx_norm), len(ocr_norm))
    combined_slides = []
    for i in range(max_slides):
        ppt = pptx_norm[i] if i < len(pptx_norm) else {"slide": i+1, "text": "", "numbers": [], "dates": []}
        ocr = ocr_norm[i] if i < len(ocr_norm) else {"slide": i+1, "text": "", "numbers": [], "dates": []}
        merged_text = (ppt["text"] + " " + ocr["text"]).strip()
        merged_numbers = list(set(ppt.get("numbers", []) + ocr.get("numbers", [])))
        merged_dates = list(set(ppt.get("dates", []) + ocr.get("dates", [])))
        combined_slides.append({
            "slide": i+1,
            "text": merged_text,
            "numbers": merged_numbers,
            "dates": merged_dates,
            "pptx_text": ppt["text"],
            "ocr_text": ocr["text"]
        })

    # Run rule-based checks
    print("[4/5] Running rule-based detectors...")
    rule_issues = []
    rule_issues.extend(detect_numeric_conflicts(combined_slides))
    rule_issues.extend(detect_percent_sum_issues(combined_slides))

    # LLM-based deep checks (if API key available)
    llm_issues = []
    api_key_present = os.environ.get("my_api_key") is not None
    mode = "deep"
    try:
        mode = "deep"  # default deep mode in this CLI
    except:
        pass

    if mode == "deep":
        if not api_key_present:
            print("⚠️  No Gemini API key found in environment variable 'my_api_key'. Skipping deep LLM checks.")
        else:
            print("[5/5] Running Gemini deep checks (batched)...")
            batch_size = 8
            for i in range(0, len(combined_slides), batch_size):
                batch = combined_slides[i:i+batch_size]
                res = call_gemini(batch)
                if isinstance(res, list):
                    llm_issues.extend(res)
                else:
                    llm_issues.append({"type":"llm_unexpected", "raw": str(res)})

    # Aggregate and save output
    out = {
        "summary": {
            "slides_processed": len(combined_slides),
            "rule_issues": len(rule_issues),
            "llm_issues": len(llm_issues)
        },
        "rule_issues": rule_issues,
        "llm_issues": llm_issues,
        "slides": [
            {"slide": s["slide"], "pptx_text": s["pptx_text"], "ocr_text": s["ocr_text"], "numbers": s["numbers"], "dates": s["dates"]}
            for s in combined_slides
        ]
    }

    out_path = Path(__file__).resolve().parent / "inconsistencies.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)

    print("\n=== Analysis Complete ===")
    print(f"Slides processed: {out['summary']['slides_processed']}")
    print(f"Rule-based issues found: {out['summary']['rule_issues']}")
    print(f"LLM issues found: {out['summary']['llm_issues']}")
    print(f"Saved report to: {out_path}")

if __name__ == "__main__":
    main()
