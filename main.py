import os, re, json, tempfile, xml.etree.ElementTree as ET
from typing import List
import fitz, openpyxl
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

app = FastAPI(title="C4 Capital FIDC API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates", "BORDERO_TEMPLATE.xlsx")

def br2f(s):
    try: return float(str(s).replace('.','').replace(',','.'))
    except: return 0.0

def _extrair_nfse(txt):
    ti = txt.find('Dados do Tomador')
    pre, tom = (txt[:ti], txt[ti:]) if ti>0 else (txt,'')
    cnpj_c = re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', pre)
    cnpj_ced = cnpj_c[-1] if cnpj_c else ''
    nomes = re.findall(r'^(.+(?:Ltda|LTDA|S\.?A\.?|EIRELI|ME|EPP).*)$', pre, re.M|re.I)
    razao_ced = nomes[0].strip() if nomes else ''
    cnpj_sm = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2}', tom)
    cnpj_sac = re.sub(r'\D','', cnpj_sm.group()) if cnpj_sm else ''
    razao_m = re.search(r'Raz[a\u00e3]o Social\s+(.+)', tom)
    razao_sac = razao_m.group(1).strip() if razao_m else ''
    cep_m = re.search(r'CEP\s+([\d]{2}\.?[\d]{3}-[\d]{3})', tom)
    cep = cep_m.group(1).replace('.','') if cep_m else ''
    email_m = re.search(r'[\w.%+-]+@[\w.-]+\.[a-zA-Z]{2,}', tom)
    email = re.sub(r'\s+','', email_m.group()) if email_m else ''
    tel_m = re.search(r'\(\d{2}\)\s*\d[\d\s-]{6,}', tom)
    tel = tel_m.group().strip() if tel_m else ''
    num_m = re.search(r'N[u\u00fa]mero da Nota Fiscal\s+(\d+)', txt)
    num_nf = num_m.group(1) if num_m else ''
    dt_m = re.search(r'Data de Emiss[a\u00e3]o\s+(\d{2}/\d{2}/\d{4})', txt)
    dt = dt_m.group(1) if dt_m else ''
    vl_all = re.findall(r'Valor L[\u00ed\u00ec]quido\s+R\$\s+([\d.,]+)', txt)
    vs_m = re.search(r'Valor dos Servi[\u00e7c]os\s+R\$\s+([\d.,]+)', txt)
    valor = br2f(vl_all[-1]) if vl_all else (br2f(vs_m.group(1)) if vs_m else 0.0)
    ch44 = re.search(r'\b\d{44}\b', txt)
    auth = re.search(r'C[o\u00f3]digo de Autenticidade\s+(\S+)', txt)
    chave = ch44.group() if ch44 else (auth.group(1) if auth else '')
    return dict(tipo='nfse', cnpj_ced=cnpj_ced, razao_ced=razao_ced,
                cnpj_sac=cnpj_sac, razao_sac=razao_sac,
                num=num_nf, dt=dt, valor=valor, chave=chave, cep=cep, email=email, tel=tel)

def _extrair_fatura(txt):
    cnpj_ced_m = re.search(r'CNPJ:?\s*([\d.]+/[\d]+-\d{2})', txt)
    cnpj_ced = cnpj_ced_m.group(1) if cnpj_ced_m else ''
    linhas = txt.split('\n'); razao_ced = linhas[0].strip()
    num_m = re.search(r'FATURA DE LOCA[\u00c7C][\u00c3A]O\s*n[\u00bao]\s*(\d+)', txt, re.I)
    num_fat = num_m.group(1).lstrip('0') or '0' if num_m else ''
    dt_m = re.search(r'Emiss[a\u00e3]o:.*?(\d{1,2} de \w+ de \d{4})', txt)
    meses = {'Janeiro':'01','Fevereiro':'02','Mar\u00e7o':'03','Abril':'04','Maio':'05','Junho':'06',
             'Julho':'07','Agosto':'08','Setembro':'09','Outubro':'10','Novembro':'11','Dezembro':'12'}
    dt = ''
    if dt_m:
        m = re.match(r'(\d{1,2}) de (\w+) de (\d{4})', dt_m.group(1))
        if m:
            d, mes_nome, a = m.groups()
            dt = f"{int(d):02d}/{meses.get(mes_nome,'01')}/{a}"
    cli_m = re.search(r'Cliente:\s*(.+)', txt)
    razao_sac = cli_m.group(1).strip() if cli_m else ''
    tom_block = txt[txt.find('Cliente:'):]
    cnpj_sac_m = re.search(r'CNPJ:?\s*([\d.]+/[\d]+-\d{2})', tom_block)
    cnpj_sac = cnpj_sac_m.group(1) if cnpj_sac_m else ''
    cep_m = re.search(r'CEP:?\s*([\d-]+)', tom_block)
    cep = cep_m.group(1) if cep_m else ''
    email_m = re.search(r'[\w.%+-]+@[\w.-]+\.[a-zA-Z]{2,}', tom_block)
    email = email_m.group(0) if email_m else ''
    tel_m = re.search(r'Telefone:?\s*(\(\d{2}\)\s*[\d-]+)', tom_block)
    tel = tel_m.group(1) if tel_m else ''
    parc_idx = txt.find('Parcela')
    bloco = txt[parc_idx:] if parc_idx>0 else ''
    parcelas_n = re.findall(r'(\d{3}/\d{3})', bloco)
    datas = re.findall(r'(\d{2}/\d{2}/\d{4})', bloco)
    valores_raw = re.findall(r'([\d.]+,\d{2})', bloco)
    n = len(parcelas_n)
    parcelas = [dict(parcela=parcelas_n[i], venc=datas[i] if i<len(datas) else '', valor=br2f(valores_raw[i]) if i<len(valores_raw) else 0) for i in range(n)]
    return dict(tipo='fatura', cnpj_ced=cnpj_ced, razao_ced=razao_ced.strip(),
                cnpj_sac=cnpj_sac, razao_sac=razao_sac,
                num_fat=num_fat, dt=dt, cep=cep, email=email, tel=tel, parcelas=parcelas)

def extrair_pdf(caminho):
    doc = fitz.open(caminho); txt = doc[0].get_text()
    if 'Fatura de Loca' in txt or 'FATURA DE LOCACAO' in txt.upper():
        return _extrair_fatura(txt)
    return _extrair_nfse(txt)

def extrair_xml(caminho):
    ns = 'http://www.portalfiscal.inf.br/nfe'
    tree = ET.parse(caminho); root = tree.getroot()
    def ft(el, tag, d=''):
        r = el.find(f'{{{ns}}}{tag}') if el is not None else None
        return r.text if r is not None else d
    ide = root.find(f'.//{{{ns}}}ide')
    emit = root.find(f'.//{{{ns}}}emit')
    dest = root.find(f'.//{{{ns}}}dest')
    total = root.find(f'.//{{{ns}}}total/{{{ns}}}ICMSTot')
    num_nf = ft(ide,'nNF'); dt_raw = ft(ide,'dhEmi')
    dt = ''
    if dt_raw:
        p = dt_raw[:10].split('-')
        dt = f"{p[2]}/{p[1]}/{p[0]}"
    cnpj_ced = ft(emit,'CNPJ'); razao_ced = ft(emit,'xNome')
    cnpj_sac = ft(dest,'CNPJ') or ft(dest,'CPF'); razao_sac = ft(dest,'xNome')
    email_sac = ft(dest,'email')
    ender = dest.find(f'{{{ns}}}enderDest') if dest else None
    cep = ft(ender,'CEP')
    valor = br2f(ft(total,'vNF'))
    chave_el = root.find(f'.//{{{ns}}}chNFe')
    chave = chave_el.text if chave_el is not None else ''
    return dict(tipo='nfe_xml', cnpj_ced=cnpj_ced, razao_ced=razao_ced,
                cnpj_sac=cnpj_sac, razao_sac=razao_sac,
                num=num_nf, dt=dt, valor=valor, chave=chave, cep=cep, email=email_sac, tel='')

def preencher_bordero(dados_list):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template nao encontrado: {TEMPLATE_PATH}")
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb['Bordero']
    for row in range(9,109):
        for col in range(2,12):
            ws.cell(row=row, column=col).value = None
    if dados_list:
        p0 = dados_list[0]
        ws['C3'] = p0.get('cnpj_ced',''); ws['C4'] = p0.get('razao_ced','')
    linha = 9
    for d in dados_list:
        if d.get('tipo') == 'fatura':
            for p in d.get('parcelas',[]):
                if linha>108: break
                pn = p['parcela'].split('/')[0]
                ws.cell(row=linha,column=2).value = d.get('cnpj_sac','')
                ws.cell(row=linha,column=3).value = d.get('razao_sac','')
                ws.cell(row=linha,column=4).value = f"{d.get('num_fat','')}-{pn}"
                ws.cell(row=linha,column=5).value = d.get('dt','')
                ws.cell(row=linha,column=6).value = p.get('venc','')
                ws.cell(row=linha,column=7).value = d.get('email','')
                ws.cell(row=linha,column=8).value = d.get('tel','')
                ws.cell(row=linha,column=9).value = p.get('valor',0)
                ws.cell(row=linha,column=11).value = d.get('cep','')
                linha += 1
        else:
            if linha>108: break
            num = d.get('num','')
            ws.cell(row=linha,column=2).value = d.get('cnpj_sac','')
            ws.cell(row=linha,column=3).value = d.get('razao_sac','')
            ws.cell(row=linha,column=4).value = f"{num}-001" if num else ''
            ws.cell(row=linha,column=5).value = d.get('dt','')
            ws.cell(row=linha,column=7).value = d.get('email','')
            ws.cell(row=linha,column=8).value = d.get('tel','')
            ws.cell(row=linha,column=9).value = d.get('valor',0)
            ws.cell(row=linha,column=10).value = d.get('chave','')
            ws.cell(row=linha,column=11).value = d.get('cep','')
            linha += 1
    out = tempfile.mktemp(suffix='.xlsx')
    wb.save(out); return out

@app.get("/")
def health(): return {"status":"ok","service":"C4 Capital FIDC API"}

@app.post("/extrair")
async def extrair(files: List[UploadFile] = File(...)):
    resultados = []
    for f in files:
        tmp = tempfile.mktemp(suffix=os.path.splitext(f.filename)[1])
        try:
            with open(tmp,'wb') as out: out.write(await f.read())
            dados = extrair_xml(tmp) if f.filename.lower().endswith('.xml') else extrair_pdf(tmp)
            resultados.append({"ok":True,"arquivo":f.filename,"dados":dados})
        except Exception as e:
            resultados.append({"ok":False,"arquivo":f.filename,"erro":str(e)})
        finally:
            if os.path.exists(tmp): os.remove(tmp)
    return {"resultados":resultados}

@app.post("/gerar-bordero")
async def gerar_bordero(files: List[UploadFile] = File(...)):
    dados_list = []; tmps = []
    try:
        for f in files:
            tmp = tempfile.mktemp(suffix=os.path.splitext(f.filename)[1])
            tmps.append(tmp)
            with open(tmp,'wb') as out: out.write(await f.read())
            dados = extrair_xml(tmp) if f.filename.lower().endswith('.xml') else extrair_pdf(tmp)
            dados_list.append(dados)
        out_path = preencher_bordero(dados_list)
        return FileResponse(out_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename='Bordero_C4Capital.xlsx')
    finally:
        for t in tmps:
            if os.path.exists(t): os.remove(t)
