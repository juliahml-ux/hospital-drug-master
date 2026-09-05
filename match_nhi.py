import csv
import re
import json
import openpyxl
import sys

def normalize_text(text):
    if not text:
        return ""
    text = str(text).lower()
    # Replace full-width symbols with half-width
    text = text.replace("（", "(").replace("）", ")").replace("＆", "&").replace("％", "%")
    # Clean quotes and unwanted symbols
    text = re.sub(r'[\"\'`]+', '', text)
    return text.strip()

def extract_tokens(text):
    norm = normalize_text(text)
    # Extract alphanumeric words and strength patterns like 500mg, 10%, 5mg/ml
    tokens = set(re.findall(r'[a-z0-9\.]+(?:mg|mcg|gm|g|ml|iu|%)?', norm))
    # Filter out pure noise tokens
    tokens = {t for t in tokens if len(t) > 1 and t not in {'and', 'the', 'for', 'with', 'in', 'of', 'tab', 'cap', 'inj', 'sol', 'vial', 'amp', 'box', 'btl', 'tube'}}
    return tokens

def extract_strength_number(text):
    if not text:
        return []
    # Find all dosages like 500mg, 20mg, 0.45%, 10%, 1g, 5g
    matches = re.findall(r'(\d+(?:\.\d+)?)\s*(mg|mcg|gm|g|ml|iu|%)', str(text).lower())
    return [f"{num}{unit}" for num, unit in matches]

def main():
    print("Step 1: Loading 761 Hospital Drugs from Excel...")
    wb = openpyxl.load_workbook('115年第一季院內藥品清單.xlsx', data_only=True)
    ws = wb.active
    hosp_drugs = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        hcode = str(r[1]).strip() if r[1] is not None else ''
        bname = str(r[2]).strip() if r[2] is not None else ''
        gname = str(r[3]).strip() if r[3] is not None else ''
        form = str(r[4]).strip() if r[4] is not None else ''
        spec = str(r[5]).strip() if r[5] is not None else ''
        unit = str(r[6]).strip() if r[6] is not None else ''
        atc = str(r[7]).strip() if r[7] is not None else ''
        if not hcode and not bname:
            continue
        full_spec = f"{spec} ({unit})" if spec and unit else (spec or unit)
        hosp_drugs.append({
            'idx': i,
            'hospitalCode': hcode,
            'brandName': bname,
            'genericName': gname,
            'strength': full_spec,
            'dosageForm': form,
            'atc': atc,
            'b_tokens': extract_tokens(bname),
            'g_tokens': extract_tokens(gname),
            'strengths': extract_strength_number(spec) + extract_strength_number(bname)
        })
    print(f"Loaded {len(hosp_drugs)} hospital drugs.")

    print("\nStep 2: Indexing NHI Master Database (output/nhi_drugs_full.csv)...")
    nhi_latest = {}
    with open('output/nhi_drugs_full.csv', 'r', encoding='utf-8-sig', errors='ignore') as f:
        reader = csv.DictReader(f)
        for row in reader:
            code = (row.get('藥品代號') or '').strip()
            if not code or len(code) != 10:
                continue
            start_date = (row.get('有效起日') or '').strip()
            price = (row.get('支付價') or '').strip()
            
            # Keep latest price and valid record
            if code not in nhi_latest or start_date >= nhi_latest[code]['startDate']:
                nhi_latest[code] = {
                    'nhiCode': code,
                    'brandName': (row.get('藥品英文名稱') or '').strip(),
                    'chineseName': (row.get('藥品中文名稱') or '').strip(),
                    'genericName': (row.get('成分') or '').strip(),
                    'specAmount': (row.get('規格量') or '').strip(),
                    'specUnit': (row.get('規格單位') or '').strip(),
                    'form': (row.get('劑型') or '').strip(),
                    'price': price,
                    'startDate': start_date,
                    'endDate': (row.get('有效迄日') or '').strip(),
                    'manufacturer': (row.get('製造廠名稱') or '').strip(),
                    'atc': (row.get('ATC代碼') or '').strip(),
                    'ruleSection': (row.get('給付規定章節') or '').strip(),
                    'ruleLink': (row.get('給付規定章節連結') or '').strip(),
                    'nhiLink': (row.get('藥品代碼超連結') or '').strip()
                }

    print(f"Indexed {len(nhi_latest)} unique active NHI drug items.")

    # Pre-process NHI token indices for fast lookup
    nhi_records = []
    atc_index = {}
    for code, rec in nhi_latest.items():
        b_tokens = extract_tokens(rec['brandName'])
        g_tokens = extract_tokens(rec['genericName'])
        strengths = extract_strength_number(rec['specAmount'] + rec['specUnit']) + extract_strength_number(rec['brandName']) + extract_strength_number(rec['genericName'])
        rec_data = {
            'rec': rec,
            'b_tokens': b_tokens,
            'g_tokens': g_tokens,
            'strengths': strengths,
            'norm_brand': normalize_text(rec['brandName'])
        }
        nhi_records.append(rec_data)
        if rec['atc']:
            atc_index.setdefault(rec['atc'].upper(), []).append(rec_data)

    print("\nStep 3: Matching Hospital Drugs against NHI Database...")
    matched_drugs = []
    matched_count = 0

    for hd in hosp_drugs:
        best_candidate = None
        best_score = 0
        
        # Candidates: prioritize same ATC if available
        candidates = atc_index.get(hd['atc'].upper(), []) if hd['atc'] else []
        if not candidates or len(candidates) < 5:
            candidates = candidates + nhi_records

        hd_norm_b = normalize_text(hd['brandName'])

        for cand in candidates:
            score = 0
            cand_rec = cand['rec']
            
            # Exact brand match
            if hd_norm_b and hd_norm_b == cand['norm_brand']:
                score += 120
            
            # Token overlap in Brand Name
            b_overlap = hd['b_tokens'].intersection(cand['b_tokens'])
            score += len(b_overlap) * 25

            # Token overlap in Generic Name
            g_overlap = hd['g_tokens'].intersection(cand['g_tokens'])
            score += len(g_overlap) * 15

            # ATC match
            if hd['atc'] and cand_rec['atc'] and hd['atc'].upper() == cand_rec['atc'].upper():
                score += 35
            elif hd['atc'] and cand_rec['atc'] and hd['atc'][:4].upper() == cand_rec['atc'][:4].upper():
                score += 15

            # Strength / Dosage match
            if hd['strengths'] and cand['strengths']:
                st_overlap = set(hd['strengths']).intersection(set(cand['strengths']))
                if st_overlap:
                    score += 30
                else:
                    score -= 15

            if score > best_score:
                best_score = score
                best_candidate = cand_rec

        # Threshold for acceptable match
        if best_score >= 45 and best_candidate:
            matched_count += 1
            nhi_code = best_candidate['nhiCode']
            chinese_name = best_candidate['chineseName']
            nhi_price = best_candidate['price']
            rule_section = best_candidate['ruleSection']
            rule_link = best_candidate['ruleLink']
            nhi_link = best_candidate['nhiLink']
            # Fallback official NHI search query URL if no direct link
            if not nhi_link:
                nhi_link = f"https://info.nhi.gov.tw/IODE0000/IODE0000S06"
        else:
            nhi_code = ""
            chinese_name = ""
            nhi_price = ""
            rule_section = ""
            rule_link = ""
            nhi_link = "https://info.nhi.gov.tw/IODE0000/IODE0000S06"

        # Construct enriched drug object
        matched_drugs.append({
            'id': f"d_115_{hd['idx']}",
            'hospitalCode': hd['hospitalCode'],
            'nhiCode': nhi_code,
            'atc': hd['atc'] or (best_candidate['atc'] if best_candidate else ''),
            'brandName': hd['brandName'],
            'chineseName': chinese_name,
            'genericName': hd['genericName'],
            'strength': hd['strength'],
            'dosageForm': hd['dosageForm'],
            'price': nhi_price,
            'ruleSection': rule_section,
            'ruleLink': rule_link,
            'nhiLink': nhi_link
        })

    print(f"\nMatching Results: {matched_count} / {len(hosp_drugs)} ({matched_count/len(hosp_drugs)*100:.1f}%) successfully matched with NHI Code & Prices!")

    # Save to JSON
    output_json = 'output/115_drugs_enriched_preset.json'
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(matched_drugs, f, ensure_ascii=False, indent=2)
    print(f"Saved enriched data to {output_json}")

    # Output Sample
    print("\n--- Top 10 Matched Samples ---")
    for d in matched_drugs[:10]:
        print(f"[{d['hospitalCode']}] {d['brandName']}")
        print(f"  -> NHI: {d['nhiCode']} | 中文: {d['chineseName']} | 健保價: NT${d['price']} | ATC: {d['atc']}")
        print(f"  -> 給付規定: {d['ruleSection']} | 規定連結: {d['ruleLink']}")

if __name__ == '__main__':
    main()
