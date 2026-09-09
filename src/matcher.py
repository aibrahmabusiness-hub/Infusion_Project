import re
import os
from difflib import SequenceMatcher

# Dictionary of common synonyms, brand names, and drug mappings
DRUG_SYNONYMS = {
    'versed': 'midazolam',
    'midazolam': 'midazolam',
    'dilaudid': 'hydromorphone',
    'hydromorphone': 'hydromorphone',
    'levophed': 'norepinephrine',
    'norepinephrine': 'norepinephrine',
    'ketamine': 'ketamine',
    'fentanyl': 'fentanyl',
    'morphine': 'morphine',
    'sublimaze': 'fentanyl',
    'demerol': 'meperidine',
    'meperidine': 'meperidine'
}

def extract_drug_keywords(text):
    """
    Extracts standardized drug names and active ingredients from description text.
    """
    if not isinstance(text, str):
        return set()
        
    text_lower = text.lower()
    
    # Strip common prefixes/suffixes and punctuation
    cleaned_text = re.sub(r'[^\w\s]', ' ', text_lower)
    tokens = cleaned_text.split()
    
    found_drugs = set()
    for token in tokens:
        # Check direct synonyms
        if token in DRUG_SYNONYMS:
            found_drugs.add(DRUG_SYNONYMS[token])
            
    # Fallback to substring matching for compound words like "fentanyl-ns" or "ketaminepca"
    for word, standardized in DRUG_SYNONYMS.items():
        if word in text_lower:
            found_drugs.add(standardized)
            
    return found_drugs

def match_medications_deterministic(desc1, desc2):
    """
    Deterministically matches two medication descriptions by comparing extracted active ingredients.
    Returns (matched, confidence, method, explanation)
    """
    if not isinstance(desc1, str) or not isinstance(desc2, str):
        return False, 0.0, "Deterministic Match", "Invalid input descriptions"
        
    drugs1 = extract_drug_keywords(desc1)
    drugs2 = extract_drug_keywords(desc2)
    
    # If we found matching active ingredients
    common_drugs = drugs1.intersection(drugs2)
    if common_drugs:
        matched_drug = list(common_drugs)[0]
        return True, 1.0, "Deterministic Match", f"Matched on active ingredient: {matched_drug.upper()}"
        
    # If no exact drug matches, compute string similarity on normalized strings as fallback
    # Remove volumes, concentrations, and formatting characters to compare just names
    def clean_for_similarity(s):
        s = re.sub(r'\d+\s*(mg|mcg|mCg|ml|mL|g)/?\d*\s*(ml|mL)?', '', s, flags=re.IGNORECASE) # concentration
        s = re.sub(r'\(\d+\s*(ml|mL)\)', '', s, flags=re.IGNORECASE) # total volume
        s = re.sub(r'[^\w\s]', ' ', s)
        s = s.lower()
        return " ".join(s.split())
        
    s1_clean = clean_for_similarity(desc1)
    s2_clean = clean_for_similarity(desc2)
    
    ratio = SequenceMatcher(None, s1_clean, s2_clean).ratio()
    
    if ratio > 0.75:
        return True, ratio, "Deterministic Similarity Match", f"High similarity score ({ratio:.2f}) on cleaned names: '{s1_clean}' and '{s2_clean}'"
        
    return False, ratio, "Deterministic Match", f"No drug overlap. Similarity: {ratio:.2f}"

def match_medications_ai(desc1, desc2):
    """
    Uses Google Generative AI (Gemini) to determine if two medication descriptions match.
    If the API call fails or is unconfigured, it falls back to the deterministic matcher.
    Returns (matched, confidence, method, explanation)
    """
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        # Fallback to deterministic matcher silently if API key is not set
        return match_medications_deterministic(desc1, desc2)
        
    try:
        import google.generativeai as genai
        import json
        
        genai.configure(api_key=api_key)
        # Using the standard gemini model
        model = genai.GenerativeModel('gemini-2.5-flash')
        
        prompt = f"""
You are a pharmacy data reconciliation assistant. Determine if the following two medication descriptions refer to the same medication/active ingredient.
Description 1 (Pyxis Dispensing System): "{desc1}"
Description 2 (Epic EMR Flowsheet or Admin Report): "{desc2}"

Respond ONLY with a JSON object containing these keys:
- "matched": boolean (true/false)
- "confidence": float (between 0.0 and 1.0)
- "explanation": string (brief explanation of the match, including any concentration or synonym mapping details)

Ensure the response contains nothing but valid JSON.
"""
        response = model.generate_content(prompt)
        text = response.text.strip()
        
        # Clean potential markdown wrapping
        if text.startswith("```json"):
            text = text[7:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()
        
        res = json.loads(text)
        return (
            bool(res.get("matched", False)),
            float(res.get("confidence", 0.0)),
            "AI-Assisted Match",
            str(res.get("explanation", "Matched by AI"))
        )
    except Exception as e:
        # Fallback to deterministic match if error occurs
        matched, conf, method, exp = match_medications_deterministic(desc1, desc2)
        return matched, conf, f"Deterministic Match (AI Failed: {str(e)})", exp

def match_medications(desc1, desc2, threshold=0.85):
    """
    Main matching function. Calls AI matcher if configured, else deterministic.
    Enforces the threshold on confidence.
    """
    # If they are exactly equal normalized, match immediately
    if isinstance(desc1, str) and isinstance(desc2, str) and desc1.strip().lower() == desc2.strip().lower():
        return True, 1.0, "Deterministic Exact Match", "Exact string match"
        
    # Run matching
    matched, confidence, method, explanation = match_medications_ai(desc1, desc2)
    
    # Enforce threshold
    if matched and confidence >= threshold:
        return True, confidence, method, explanation
    elif matched and confidence < threshold:
        return False, confidence, method, f"Match found but confidence ({confidence:.2f}) is below threshold ({threshold:.2f}). Reason: {explanation}"
    else:
        return False, confidence, method, f"No match. Reason: {explanation}"
