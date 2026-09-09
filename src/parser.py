import re
import math

def parse_amount_string(amount_str):
    """
    Parses strings like '250 mg', '50 mg', '1250 mCg', '500 mCg', '136 mg'
    Returns a tuple of (value_float, unit_str_lowercase)
    """
    if not isinstance(amount_str, str):
        if isinstance(amount_str, (int, float)) and not math.isnan(amount_str):
            return float(amount_str), None
        return None, None
    
    amount_str = amount_str.strip()
    match = re.match(r"^(\d+(?:\.\d+)?)\s*(mg|mcg|mCg|g|mL|ml|ML)$", amount_str, re.IGNORECASE)
    if match:
        val = float(match.group(1))
        unit = match.group(2).lower()
        if unit == 'mcg' or unit == 'mcg':
            unit = 'mcg'
        return val, unit
    return None, None

def parse_med_description(desc):
    """
    Parses a Pyxis MedDescription to extract drug name, concentration, volume, and unit.
    Examples:
      - 'ketamine in NS (Ketamine) 5 mg/1 mL (50 mL) Syringe'
      - 'midazolam in NS (Versed Generic) 1 mg/1 mL (50 mL) Syringe'
      - 'fentaNYL (fentaNYL) 50 mCg/1 mL (25 mL) Syringe'
      - 'midazolam (midazolam) 50 mg (50 mL) Injection'
    
    Returns a dict with:
      - drug_name (lowercase, e.g. 'ketamine', 'midazolam', 'fentanyl')
      - concentration_strength (float, e.g. 5.0)
      - concentration_unit (lowercase, e.g. 'mg' or 'mcg')
      - concentration_volume (float, e.g. 1.0)
      - total_volume (float, e.g. 50.0)
      - concentration_per_ml (float, e.g. 5.0 unit/mL)
      - calculated_total_amount (float, e.g. 250.0)
      - calculated_total_unit (lowercase, e.g. 'mg' or 'mcg')
    """
    if not isinstance(desc, str):
        return None
        
    # Extract core drug name
    drug_name = "unknown"
    desc_lower = desc.lower()
    if "ketamine" in desc_lower:
        drug_name = "ketamine"
    elif "midazolam" in desc_lower or "versed" in desc_lower:
        drug_name = "midazolam"
    elif "fentanyl" in desc_lower:
        drug_name = "fentanyl"
    elif "morphine" in desc_lower:
        drug_name = "morphine"
    elif "hydromorphone" in desc_lower:
        drug_name = "hydromorphone"
    elif "milrinone" in desc_lower:
        drug_name = "milrinone"
    elif "dexmedetomidine" in desc_lower:
        drug_name = "dexmedetomidine"
        
    # Regex 1: Matches pattern like "5 mg/1 mL (50 mL)" or "10 mCg/1 mL (50 mL)"
    m1 = re.search(r"(\d+(?:\.\d+)?)\s*(mg|mcg|mCg)/(\d+(?:\.\d+)?)\s*(mL|ml|ML)\s*\((\d+(?:\.\d+)?)\s*(mL|ml|ML)\)", desc, re.IGNORECASE)
    if m1:
        conc_val = float(m1.group(1))
        conc_unit = m1.group(2).lower()
        conc_vol = float(m1.group(3))
        tot_vol = float(m1.group(5))
        
        concentration = conc_val / conc_vol
        total_med = concentration * tot_vol
        return {
            'drug_name': drug_name,
            'type': 'concentration_and_volume',
            'concentration_strength': conc_val,
            'concentration_unit': conc_unit,
            'concentration_volume': conc_vol,
            'total_volume': tot_vol,
            'concentration_per_ml': concentration,
            'calculated_total_amount': total_med,
            'calculated_total_unit': conc_unit
        }
        
    # Regex 2: Matches pattern like "50 mg (50 mL)" without the slash concentration
    m2 = re.search(r"(\d+(?:\.\d+)?)\s*(mg|mcg|mCg)\s*\((\d+(?:\.\d+)?)\s*(mL|ml|ML)\)", desc, re.IGNORECASE)
    if m2:
        tot_val = float(m2.group(1))
        tot_unit = m2.group(2).lower()
        tot_vol = float(m2.group(3))
        
        concentration = tot_val / tot_vol
        return {
            'drug_name': drug_name,
            'type': 'total_amount_and_volume',
            'concentration_strength': tot_val,
            'concentration_unit': tot_unit,
            'concentration_volume': tot_vol,
            'total_volume': tot_vol,
            'concentration_per_ml': concentration,
            'calculated_total_amount': tot_val,
            'calculated_total_unit': tot_unit
        }
        
    return {
        'drug_name': drug_name,
        'type': 'unknown',
        'concentration_strength': None,
        'concentration_unit': None,
        'concentration_volume': None,
        'total_volume': None,
        'concentration_per_ml': None,
        'calculated_total_amount': None,
        'calculated_total_unit': None
    }

def convert_mass_to_volume(mass_val, mass_unit, conc_per_ml, conc_unit):
    """
    Converts a mass (e.g. 136 mg) to volume (mL) using the concentration.
    Returns None if inputs are invalid.
    """
    if mass_val is None or conc_per_ml is None or conc_per_ml == 0:
        return None
        
    mass_unit_clean = mass_unit.lower() if isinstance(mass_unit, str) else ""
    conc_unit_clean = conc_unit.lower() if isinstance(conc_unit, str) else ""
    
    # Normalize units: if one is mcg and other is mg
    factor = 1.0
    if mass_unit_clean == 'mcg' and conc_unit_clean == 'mg':
        # mass is in mcg, conc is in mg/mL (1 mcg = 0.001 mg)
        factor = 0.001
    elif mass_unit_clean == 'mg' and conc_unit_clean == 'mcg':
        # mass is in mg, conc is in mcg/mL (1 mg = 1000 mcg)
        factor = 1000.0
        
    normalized_mass = mass_val * factor
    return normalized_mass / conc_per_ml

def convert_volume_to_mass(vol_val_ml, conc_per_ml, conc_unit):
    """
    Converts a volume (mL) to mass (mg/mcg) using the concentration.
    Returns (mass_val, mass_unit).
    """
    if vol_val_ml is None or conc_per_ml is None:
        return None, None
        
    mass_val = vol_val_ml * conc_per_ml
    return mass_val, conc_unit
