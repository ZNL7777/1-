import streamlit as st
import pandas as pd
import json
import uuid
import time
import re
import copy
from datetime import datetime, timedelta

# --- 页面配置 ---
st.set_page_config(
    page_title="IATF 审计转换工具 (v54.2 EMS修复版)",
    page_icon="🛡️",
    layout="wide"
)

# --- 1. 侧边栏：模板加载 ---
with st.sidebar:
    st.header("⚙️ 模板配置")
    
    st.info("💡 请上传您的 JSON 模板。程序将把该文件作为完整的底层骨架。")
    user_template_file = st.file_uploader("上传基础 JSON 模板", type=["json"])
    
    base_template_data = None
    if user_template_file:
        try:
            base_template_data = json.load(user_template_file)
            st.success(f"✅ 成功加载底座模板: {user_template_file.name}")
        except Exception as e:
            st.error(f"❌ 模板解析失败: {e}")
            st.stop()
    else:
        st.warning("👈 请先在左侧上传 JSON 模板文件。")
        st.stop()

# --- 辅助函数：安全寻址 ---
def ensure_path(d, path):
    current = d
    for key in path:
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    return current

def safe_get(obj, key, default=""):
    if isinstance(obj, dict):
        return obj.get(key, default)
    return default

# --- 核心转换逻辑 ---
def generate_json_logic(excel_file, base_data):
    # 💥 直接使用用户上传的模板作为最终底座
    final_json = copy.deepcopy(base_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
        doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8], header=None) if len(xls.sheet_names) >= 9 else pd.DataFrame()
    except Exception as e:
        raise ValueError(f"Excel 读取失败: {str(e)}")

    def find_val_by_key(df, keywords, col_offset=1):
        if df.empty: return ""
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_val = str(df.iloc[r, c]).strip()
                for k in keywords:
                    if k in cell_val:
                        if c + col_offset < df.shape[1]:
                            return str(df.iloc[r, c + col_offset]).strip()
        return ""
        
    def get_db_val(r, c):
        try:
            val = db_df.iloc[r, c]
            return str(val).strip() if pd.notna(val) else ""
        except: return ""

    # ================= 2. 数据提取 =================
    
    # [辅助函数：暴力剥离所有非英文字母，并重排大小写]
    def extract_and_format_english_name(raw_val):
        clean_val = str(raw_val).replace("姓名:", "").replace("Name:", "").strip()
        if not clean_val: return ""
        
        eng_only = re.sub(r'[^a-zA-Z\s]', ' ', clean_val).strip()
        eng_only = re.sub(r'\s+', ' ', eng_only)
        
        if eng_only:
            parts = eng_only.split()
            if len(parts) >= 2 and parts[0].isupper() and not parts[1].isupper():
                return f"{parts[1]} {parts[0]}"
            else:
                return eng_only
        return clean_val

    # [全局姓名：保持原始中文不处理，供 AuditData 使用]
    raw_name_full = find_val_by_key(db_df, ["姓名", "Auditor Name"]) or get_db_val(5, 1)
    raw_name = raw_name_full.replace("姓名:", "").replace("Name:", "").strip() if raw_name_full else ""

    # [生成格式化后的英文名，专门供 AuditTeam 的 Name 使用]
    formatted_team_name = extract_and_format_english_name(raw_name_full)

    # [CCAA]
    ccaa_raw = find_val_by_key(db_df, ["审核员CCAA", "CCAA"]) or get_db_val(4, 1)
    caa_no = ""
    if ccaa_raw:
        match = re.search(r'(?:CCAA[:：\s-])\s*(.*)', ccaa_raw, re.IGNORECASE | re.DOTALL)
        caa_no = match.group(1).strip() if match else ccaa_raw.strip()

    # [AuditorId]
    auditor_id = ""
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                cell_text = str(info_df.iloc[r, c])
                if "IATF Card" in cell_text or "IATF卡号" in cell_text:
                    if c + 1 < info_df.shape[1]:
                        raw_val = str(info_df.iloc[r, c + 1]).strip()
                        raw_val = raw_val.replace('\n', ' ').replace('\r', ' ')
                        auditor_id = re.sub(r'^IATF[:：\s-]*', '', raw_val, flags=re.IGNORECASE).strip()
                        if len(auditor_id) > 4: break
            if auditor_id and len(auditor_id) > 4: break

    # [日期]
    start_date_raw = find_val_by_key(db_df, ["审核开始日期", "审核开始时间"]) or get_db_val(2, 1)
    end_date_raw = find_val_by_key(db_df, ["审核结束日期", "审核结束时间"]) or get_db_val(3, 1)
    
    def fmt_iso(val):
        try:
            clean_val = str(val).replace('年', '-').replace('月', '-').replace('日', '')
            dt = pd.to_datetime(clean_val, errors='coerce')
            if pd.notna(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""
        
    start_iso, end_iso = fmt_iso(start_date_raw), fmt_iso(end_date_raw)
    
    next_audit_iso = ""
    try:
        clean_end = str(end_date_raw).replace('年', '-').replace('月', '-').replace('日', '')
        end_dt = pd.to_datetime(clean_end, errors='coerce')
        if pd.notna(end_dt): next_audit_iso = (end_dt + timedelta(days=45)).strftime('%Y-%m-%d') + "T00:00:00.000Z"
    except: pass

    # [多顾客与 CSR 动态提取]
    customers_list = []
    if not info_df.empty:
        header_r = -1
        col_map = {'cust': -1, 'name': -1, 'date': -1, 'code': -1}
        for r in range(info_df.shape[0]):
            row_str = " ".join([str(x) for x in info_df.iloc[r, :]]).upper()
            if "CUSTOMER" in row_str and ("CSR" in row_str or "TITLE" in row_str):
                header_r = r
                for c in range(info_df.shape[1]):
                    val = str(info_df.iloc[r, c]).strip().upper()
                    if "CUSTOMER" in val or "客户" in val: col_map['cust'] = c
                    elif "CSR" in val or "TITLE" in val: col_map['name'] = c
                    elif "VERSION" in val or "DATE" in val or "版本" in val or "日期" in val: col_map['date'] = c
                    elif "供应商代码" in val or "SUPPLIER" in val or "CODE" in val: col_map['code'] = c
                break
                
        if header_r != -1:
            for r in range(header_r + 1, info_df.shape[0]):
                cust_val = str(info_df.iloc[r, col_map['cust']]).strip() if col_map['cust'] != -1 else ""
                if not cust_val or cust_val.lower() == 'nan': continue
                if "审核员" in cust_val or "AUDIT" in cust_val.upper() or "NAME" in cust_val.upper(): break
                    
                name_val = str(info_df.iloc[r, col_map['name']]).strip() if col_map['name'] != -1 else ""
                date_val = str(info_df.iloc[r, col_map['date']]).strip() if col_map['date'] != -1 else ""
                code_val = str(info_df.iloc[r, col_map['code']]).strip() if col_map['code'] != -1 else ""
                
                if name_val.lower() == 'nan': name_val = ""
                if date_val.lower() == 'nan': date_val = ""
                if code_val.lower() == 'nan': code_val = ""

                final_date = date_val.replace(" 00:00:00", "").strip()

                customers_list.append({
                    "Name": cust_val,
                    "SupplierCode": code_val,
                    "NameCSRDocument": name_val,
                    "DateCSRDocument": final_date
                })

    if not customers_list:
        customer_name = find_val_by_key(db_df, ["顾客", "客户名称"]) or get_db_val(29, 1)
        supplier_code = find_val_by_key(db_df, ["供应商编码", "供应商代码"]) or get_db_val(30, 1)
        csr_name = find_val_by_key(db_df, ["CSR文件名称"]) or get_db_val(31, 1)
        csr_date_raw = find_val_by_key(db_df, ["CSR文件日期"]) or get_db_val(32, 1)
        
        csr_date = str(csr_date_raw).replace(" 00:00:00", "").strip()
        if csr_date.lower() == 'nan': csr_date = ""
        
        if customer_name or supplier_code or csr_name:
            customers_list.append({
                "Name": customer_name,
                "SupplierCode": supplier_code,
                "NameCSRDocument": csr_name,
                "DateCSRDocument": csr_date
            })

    # 💥 [新增：EMS扩展场所 动态全量提取与地址切分]
    ems_sites = []
    if not info_df.empty:
        header_r = -1
        col_map = {}
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                val = str(info_df.iloc[r, c]).strip().upper()
                # 兼容多种常见的 EMS 表头名称
                if "EMS扩展场所信息" in val or "扩展制造场所" in val or "扩展现场" in val:
                    header_r = r
                    for c_scan in range(info_df.shape[1]):
                        h_val = str(info_df.iloc[r, c_scan]).strip()
                        if "中文名称" in h_val: col_map['name_cn'] = c_scan
                        elif "英文名称" in h_val: col_map['name_en'] = c_scan
                        elif "中文地址" in h_val: col_map['addr_cn'] = c_scan
                        elif "英文地址" in h_val: col_map['addr_en'] = c_scan
                        elif "邮编" in h_val or "邮政编码" in h_val: col_map['zip'] = c_scan
                        elif "USI" in h_val.upper(): col_map['usi'] = c_scan
                        elif "人数" in h_val: col_map['emp'] = c_scan
                    break
            if header_r != -1: break
                
        if header_r != -1:
            for r in range(header_r + 1, info_df.shape[0]):
                def safe_get_cell(row, col_idx):
                    if col_idx == -1: return ""
                    v = str(info_df.iloc[row, col_idx]).strip()
                    return "" if v.lower() == 'nan' else v

                name_cn = safe_get_cell(r, col_map.get('name_cn', -1))
                name_en = safe_get_cell(r, col_map.get('name_en', -1))
                addr_cn = safe_get_cell(r, col_map.get('addr_cn', -1))
                
                # 如果遇到空行，跳过；如果遇到其他模块表头，跳过
                if not name_cn and not addr_cn: continue
                if "名称" in name_cn and "地址" in addr_cn: continue
                
                # 拼接中英文名称
                full_site_name = name_cn
                if name_en and name_en not in name_cn:
                    full_site_name = f"{name_cn} {name_en}".strip()

                addr_en = safe_get_cell(r, col_map.get('addr_en', -1))
                zip_code = safe_get_cell(r, col_map.get('zip', -1))
                usi = safe_get_cell(r, col_map.get('usi', -1))
                emp = safe_get_cell(r, col_map.get('emp', -1))

                # 执行倒序切分
                ems_street, ems_city, ems_state, ems_country = addr_en, "", "", ""
                if addr_en:
                    clean_eng = addr_en.replace('，', ',')
                    parts = [p.strip() for p in clean_eng.split(',') if p.strip()]
                    if len(parts) >= 3:
                        ems_country = parts[-1]
                        ems_state = parts[-2]
                        ems_city = parts[-3]
                        ems_street = ", ".join(parts[:-3])
                    else:
                        ems_street = addr_en

                site_obj = {
                    "Id": str(uuid.uuid4()),
                    "SiteName": full_site_name,
                    "IATF_USI": usi,
                    "TotalNumberEmployees": emp,
                    "AddressNative": {
                        "Street1": addr_cn,
                        "City": "",
                        "State": "",
                        "Country": "中国",
                        "PostalCode": zip_code
                    },
                    "Address": {
                        "Street1": ems_street,
                        "City": ems_city,
                        "State": ems_state,
                        "Country": ems_country,
                        "PostalCode": zip_code
                    }
                }
                ems_sites.append(site_obj)

    # [终极防弹地址扫描 (主地址)]
    english_address = ""
    native_street = ""
    
    def get_adjacent_cells(df, keywords):
        cands = []
        if df.empty: return cands
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iloc[r, c]).strip().upper()
                for k in keywords:
                    if k.upper() in val:
                        if c + 1 < df.shape[1]: cands.append(str(df.iloc[r, c+1]))
                        if c + 2 < df.shape[1]: cands.append(str(df.iloc[r, c+2]))
                        if c + 3 < df.shape[1]: cands.append(str(df.iloc[r, c+3]))
                        if c + 4 < df.shape[1]: cands.append(str(df.iloc[r, c+4]))
        return [str(x).strip() for x in cands if str(x).strip() and str(x).lower() != 'nan']

    raw_cands = get_adjacent_cells(info_df, ["审核地址", "AUDIT ADDRESS", "ADDRESS"]) + \
                get_adjacent_cells(db_df, ["地址", "ADDRESS"])
                
    en_candidates = []
    zh_candidates = []

    for cand in raw_cands:
        lines = cand.replace('\r', '\n').split('\n')
        en_parts = []
        zh_parts = []
        for line in lines:
            line = line.strip()
            if not line: continue
            if re.search(r'[\u4e00-\u9fff]', line):
                zh_parts.append(line)
            elif re.search(r'[a-zA-Z]', line):
                en_parts.append(line)
        
        if en_parts: en_candidates.append(" ".join(en_parts))
        if zh_parts: zh_candidates.append(" ".join(zh_parts))

    english_address = max(en_candidates, key=len) if en_candidates else ""
    native_street = max(zh_candidates, key=len) if zh_candidates else ""

    if not english_address:
        for df in [info_df, db_df]:
            if df.empty: continue
            for r in range(df.shape[0]):
                for c in range(df.shape[1]):
                    val = str(df.iloc[r, c]).strip()
                    val_upper = val.upper()
                    if ("PROVINCE" in val_upper or "CITY" in val_upper) and re.search(r'[a-zA-Z]', val):
                        lines = val.replace('\r', '\n').split('\n')
                        en_parts = [l.strip() for l in lines if not re.search(r'[\u4e00-\u9fff]', l) and re.search(r'[a-zA-Z]', l)]
                        if en_parts:
                            cand = " ".join(en_parts)
                            if len(cand) > len(english_address):
                                english_address = cand

    street, city, state, country = english_address, "", "", ""
    if english_address:
        clean_eng = english_address.replace('，', ',')
        parts = [p.strip() for p in clean_eng.split(',') if p.strip()]
        if len(parts) >= 3:
            country = parts[-1]
            state = parts[-2]
            city = parts[-3]
            street = ", ".join(parts[:-3])
        else:
            street = english_address

    # ================= 3. 定点替换入 final_json =================

    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # A. 审核数据
    ensure_path(final_json, ["AuditData", "AuditDate"])
    if start_iso: final_json["AuditData"]["AuditDate"]["Start"] = start_iso
    if end_iso: final_json["AuditData"]["AuditDate"]["End"] = end_iso
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"]) or get_db_val(2, 4)
    
    # 💥 AuditorName 保持原始未处理文本
    final_json["AuditData"]["AuditorName"] = raw_name
    final_json["AuditData"]["auditorname"] = raw_name

    if "AuditTeam" not in final_json["AuditData"] or not isinstance(final_json["AuditData"]["AuditTeam"], list) or len(final_json["AuditData"]["AuditTeam"]) == 0:
        final_json["AuditData"]["AuditTeam"] = [{}]
        
    team = final_json["AuditData"]["AuditTeam"][0]
    if isinstance(team, dict):
        team.update({
            "Name": formatted_team_name, # 💥 前端 Name 应用格式化英文名
            "CaaNo": caa_no,
            "AuditorId": auditor_id, 
            "AuditDaysPerformed": 1.5,
            "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}]
        })

    # B. 组织与地址信息 
    ensure_path(final_json, ["OrganizationInformation", "AddressNative"])
    ensure_path(final_json, ["OrganizationInformation", "Address"])
    org = final_json["OrganizationInformation"]
    
    org["OrganizationName"] = find_val_by_key(db_df, ["组织名称"]) or get_db_val(1, 4)
    org["IndustryCode"] = find_val_by_key(db_df, ["行业代码", "Industry Code"])
    org["IATF_USI"] = find_val_by_key(db_df, ["IATF USI", "USI"]) or get_db_val(3, 4)
    org["TotalNumberEmployees"] = find_val_by_key(db_df, ["包括扩展现场在内的员工总数", "员工总数"]) or get_db_val(27, 1)
    org["CertificateScope"] = find_val_by_key(db_df, ["证书范围"])
    org["Representative"] = find_val_by_key(db_df, ["组织代表", "管理者代表", "联系人", "Representative"]) or get_db_val(15, 1)
    org["Telephone"] = find_val_by_key(db_df, ["联系电话", "电话", "Telephone"]) or get_db_val(15, 4)
    extracted_email = find_val_by_key(db_df, ["电子邮箱", "邮箱", "Email", "E-mail"]) or get_db_val(16, 1)
    org["Email"] = "" if str(extracted_email).strip() == "0" else extracted_email
    
    if "LanguageByManufacturingPersonnel" in org:
        lang_node = org["LanguageByManufacturingPersonnel"]
        if isinstance(lang_node, list) and len(lang_node) > 0:
            if isinstance(lang_node[0], dict): lang_node[0]["Products"] = ""
        elif isinstance(lang_node, dict):
            if "0" in lang_node and isinstance(lang_node["0"], dict): lang_node["0"]["Products"] = ""
            else: lang_node["Products"] = ""
    
    org["AddressNative"].update({
        "Street1": native_street,
        "State": "", "City": "", "Country": "中国",
        "PostalCode": find_val_by_key(db_df, ["邮政编码"]) or get_db_val(10, 4)
    })
    
    org["Address"].update({
        "State": state, "City": city, "Country": country, "Street1": street,
        "PostalCode": find_val_by_key(db_df, ["邮政编码"]) or get_db_val(10, 4)
    })

    # 💥 【修改点】：将 EMS 赋予根节点，并同步更新标志位
    if ems_sites:
        final_json["ExtendedManufacturingSites"] = ems_sites
        org["ExtendedManufacturingSite"] = "1"
    else:
        org["ExtendedManufacturingSite"] = "0"

    # C. 顾客与 CSR 
    ensure_path(final_json, ["CustomerInformation"])
    final_json["CustomerInformation"]["Customers"] = []
    
    for c_info in customers_list:
        cust_obj = {
            "Id": str(uuid.uuid4()),
            "Name": c_info["Name"],
            "SupplierCode": c_info["SupplierCode"],
            "Csrs": [
                {
                    "Id": str(uuid.uuid4()), 
                    "Name": c_info["Name"], 
                    "SupplierCode": c_info["SupplierCode"],
                    "NameCSRDocument": c_info["NameCSRDocument"],
                    "DateCSRDocument": c_info["DateCSRDocument"]
                }
            ]
        }
        final_json["CustomerInformation"]["Customers"].append(cust_obj)

    # D. 文件清单定点替换
    docs_list = []
    if not doc_list_df.empty:
        for c in range(doc_list_df.shape[1]):
            for r in range(doc_list_df.shape[0]):
                cell_val = str(doc_list_df.iloc[r, c]).strip()
                if "公司内对应的程序文件" in cell_val or "包含名称、编号、版本" in cell_val:
                    for r2 in range(r + 1, doc_list_df.shape[0]):
                        val = str(doc_list_df.iloc[r2, c]).strip()
                        if val and val.lower() != 'nan':
                            docs_list.append(val)
                    break
            if docs_list: break

    if docs_list:
        ensure_path(final_json, ["Stage1DocumentedRequirements"])
        if "IatfClauseDocuments" not in final_json["Stage1DocumentedRequirements"] or not isinstance(final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"], list):
            final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"] = []
            
        clause_docs = final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"]
        for i, doc_name in enumerate(docs_list):
            if i < len(clause_docs):
                if isinstance(clause_docs[i], dict):
                    clause_docs[i]["DocumentName"] = doc_name
            else:
                clause_docs.append({"DocumentName": doc_name})

    # E. 过程清单重建
    processes = []
    if not proc_df.empty:
        clause_cols = proc_df.columns[13:] if proc_df.shape[1] > 13 else









