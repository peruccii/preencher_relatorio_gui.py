#!/usr/bin/env python3
"""
preencher_relatorio_gui.py
Versão robusta do gerador de relatórios com fallback para CLI quando o Tkinter
não estiver disponível (resolve ModuleNotFoundError: No module named 'tkinter').
Funcionalidades:
- Consulta ReceitaWS por CNPJ
- Preenche placeholders em template .docx
- Gera [OBJETIVO_EMPRESA] opcional via provedor de IA (pluggable)
- Insere hyperlink para [LINK_DRIVE] e [LINK_PARA_DOWNLOAD]
- Modo GUI (Tkinter) quando disponível; caso contrário, modo CLI automático
- Argumentos de linha de comando para rodar em modo não-GUI
- Testes unitários simples acessíveis via --run-tests
"""
from __future__ import annotations
import re
import os
import time
import sys
import json
import argparse
import requests
from pathlib import Path
from typing import Dict, Optional
# tenta importar tkinter dinamicamente (alguns ambientes não têm suporte)
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    TKINTER_AVAILABLE = True
except Exception:
    TKINTER_AVAILABLE = False
    tk = None
    filedialog = None
    messagebox = None
# docx (necessário instalar python-docx)
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
# ----------------- Configuração -----------------
RECEITAWS_URL = "https://www.receitaws.com.br/v1/cnpj/{}"
REQUEST_TIMEOUT = 10
PLACEHOLDER_PATTERN = re.compile(r'\[([A-Z0-9_]+)\]')
# ----------------- Utilitários -----------------
def normalize_cnpj(cnpj_raw: str) -> str:
    digits = re.sub(r'\D', '', cnpj_raw or '')
    if len(digits) != 14:
        raise ValueError("CNPJ deve conter 14 dígitos (após remover pontuação).")
    return digits

def consulta_empresa(cnpj: str) -> dict:
    url = RECEITAWS_URL.format(cnpj)
    tries = 0
    while tries < 3:
        tries += 1
        try:
            resp = requests.get(url, timeout=REQUEST_TIMEOUT)
            if resp.status_code != 200:
                if resp.status_code in (429, 500, 502, 503, 504):
                    time.sleep(1 + tries)
                    continue
                resp.raise_for_status()
            data = resp.json()
            if isinstance(data, dict) and data.get("status") == "ERROR":
                raise RuntimeError(f"ReceitaWS retornou erro: {data.get('message')}")
            return data
        except requests.RequestException as e:
            if tries >= 3:
                raise RuntimeError(f"Falha ao consultar ReceitaWS: {e}")
            time.sleep(1 + tries)
    raise RuntimeError("Não foi possível consultar ReceitaWS após tentativas.")

def build_mapping(data: dict) -> dict:
    def safe_get(key, default=""):
        v = data.get(key)
        if v is None:
            return default
        if isinstance(v, str):
            return v.strip()
        return str(v)

    atividade_principal = ""
    if data.get("atividade_principal"):
        try:
            atividade_principal = data["atividade_principal"][0].get("text", "")
        except Exception:
            atividade_principal = str(data.get("atividade_principal"))
    resumo = " | ".join(filter(None, [
        atividade_principal,
        safe_get("porte"),
        safe_get("situacao")
    ]))
    endereco = " - ".join(filter(None, [
        safe_get("logradouro"),
        safe_get("numero"),
        safe_get("bairro"),
        safe_get("municipio"),
        safe_get("uf"),
        safe_get("cep"),
    ]))
    mapping = {
        "NOME_EMPRESA_CLIENTE": safe_get("nome"),
        "FANTASIA": safe_get("fantasia"),
        "RESUMO_EMPRESA_CLIENTE": resumo,
        "CNPJ": safe_get("cnpj"),
        "ENDERECO": endereco,
        "ATIVIDADE_PRINCIPAL": atividade_principal,
        "TELEFONE": safe_get("telefone"),
        "EMAIL": safe_get("email"),
        "ABERTURA": safe_get("abertura"),
        "SITUACAO": safe_get("situacao"),
        "OBJETIVO_EMPRESA": "",
        "LINK_DRIVE": "",
        "LINK_DRIVE_TEXT": "",
        "LINK_PARA_DOWNLOAD": "",
        "LINK_PARA_DOWNLOAD_TEXT": "",
        "DATA_BACKUP": "",
        "DATA_KICKOFF": "",
        "DATA_ENTREGA": "",
        "DOMINIO": "",
        "DEMANDA": "",
        "DOMINIOWP": "",
        "ESPECIALISTARESPONSAVEL": "ITALO GOMES",
    }
    return mapping

# ----------------- AI Providers -----------------
class AIProviderBase:
    def generate_objective(self, source_text: str, context: dict) -> str:
        raise NotImplementedError

class MockProvider(AIProviderBase):
    def generate_objective(self, source_text: str, context: dict) -> str:
        nome = context.get("NOME_EMPRESA_CLIENTE", "").strip()
        atividade = context.get("ATIVIDADE_PRINCIPAL", "").strip()
        if atividade:
            return (
                f"O objetivo da {nome} é atuar em {atividade.lower()}, oferecendo soluções "
                "e serviços relacionados a essa atividade, com foco em qualidade e atendimento ao cliente."
            )
        if source_text:
            first_sent = source_text.split(".")[0].strip()
            if first_sent:
                return f"O objetivo da {nome} é {first_sent}."
        return f"O objetivo da {nome} é oferecer produtos/serviços no seu segmento de atuação."

class HuggingFaceProvider(AIProviderBase):
    def __init__(self, api_token: Optional[str] = None, model: str = "google/flan-t5-large"):
        self.api_token = api_token or os.environ.get("HUGGINGFACE_API_TOKEN")
        self.model = model
        if not self.api_token:
            raise RuntimeError("Hugging Face token não configurado (HUGGINGFACE_API_TOKEN).")

    def generate_objective(self, source_text: str, context: dict) -> str:
        prompt = (
            "Você é um assistente que escreve um 'Objetivo da Empresa' curto (1-2 parágrafos) "
            "baseado nas informações abaixo. Seja direto e formal.\n\n"
            f"INFORMAÇÕES:\n{source_text}\n\n"
            "RETORNE APENAS o texto final, sem rótulos."
        )
        url = f"https://api-inference.huggingface.co/models/{self.model}"
        headers = {"Authorization": f"Bearer {self.api_token}"}
        payload = {"inputs": prompt, "options": {"wait_for_model": True}}
        resp = requests.post(url, json=payload, headers=headers, timeout=30)
        if resp.status_code != 200:
            raise RuntimeError(f"HF API erro {resp.status_code}: {resp.text}")
        result = resp.json()
        text = ""
        if isinstance(result, list) and result:
            first = result[0]
            if isinstance(first, dict):
                text = first.get("generated_text") or first.get("text") or str(first)
            else:
                text = str(first)
        elif isinstance(result, dict):
            text = result.get("generated_text") or result.get("text") or json.dumps(result)
        else:
            text = str(result)
        return (text or "").strip()

class OpenAIProvider(AIProviderBase):
    def __init__(self, api_key: Optional[str] = None, model: str = "gpt-4o-mini"):
        self.api_key = api_key or os.environ.get("OPENAI_API_KEY")
        self.model = model
        if not self.api_key:
            raise RuntimeError("OPENAI_API_KEY não configurada.")
        try:
            from openai import OpenAI
            self.client = OpenAI(api_key=self.api_key)
        except Exception as e:
            raise RuntimeError("Biblioteca openai não instalada. pip install openai") from e

    def generate_objective(self, source_text: str, context: dict) -> str:
        prompt = (
            "Escreva um texto curto (1-2 parágrafos) intitulado 'Objetivo da Empresa' baseado nas "
            "informações abaixo. Use linguagem formal e direta. Retorne apenas o texto.\n\n"
            f"INFORMAÇÕES:\n{source_text}\n"
        )
        resp = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=256,
            temperature=0.2,
        )
        content = resp.choices[0].message.content.strip()
        return content

def get_ai_provider(name: Optional[str]) -> AIProviderBase:
    name = (name or "mock").lower()
    if name == "mock":
        return MockProvider()
    if name in ("hf", "huggingface"):
        return HuggingFaceProvider()
    if name in ("openai", "gpt"):
        return OpenAIProvider()
    raise ValueError(f"Provider IA desconhecido: {name}")

# ----------------- DOCX helpers -----------------
def add_hyperlink(paragraph, url: str, text: str):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), "0000FF")
    rPr.append(c)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    new_run.append(rPr)
    new_t = OxmlElement("w:t")
    new_t.text = text
    new_run.append(new_t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def replace_in_paragraph(paragraph, mapping: Dict[str, str]):
    full_text = "".join([r.text for r in paragraph.runs])
    non_link_keys = {k: v for k, v in mapping.items() if k not in ("LINK_DRIVE", "LINK_DRIVE_TEXT", "LINK_PARA_DOWNLOAD", "LINK_PARA_DOWNLOAD_TEXT")}

    # Handle [LINK_DRIVE]
    if "[LINK_DRIVE]" in full_text and mapping.get("LINK_DRIVE"):
        parts = full_text.split("[LINK_DRIVE]")
        # Clear all runs
        for i in range(len(paragraph.runs) - 1, -1, -1):
            paragraph._element.remove(paragraph.runs[i]._element)
        for idx, part in enumerate(parts):
            # Replace other placeholders in this part
            for key, val in non_link_keys.items():
                part = part.replace(f'[{key}]', val or "")
            if part:
                paragraph.add_run(part)
            if idx < len(parts) - 1:
                display = mapping.get("LINK_DRIVE_TEXT") or "Link Drive"
                add_hyperlink(paragraph, mapping["LINK_DRIVE"], display)
        return

    # Handle [LINK_PARA_DOWNLOAD]
    if "[LINK_PARA_DOWNLOAD]" in full_text and mapping.get("LINK_PARA_DOWNLOAD"):
        parts = full_text.split("[LINK_PARA_DOWNLOAD]")
        # Clear all runs
        for i in range(len(paragraph.runs) - 1, -1, -1):
            paragraph._element.remove(paragraph.runs[i]._element)
        for idx, part in enumerate(parts):
            # Replace other placeholders in this part
            for key, val in non_link_keys.items():
                part = part.replace(f'[{key}]', val or "")
            if part:
                paragraph.add_run(part)
            if idx < len(parts) - 1:
                display = "Link para download"
                add_hyperlink(paragraph, mapping["LINK_PARA_DOWNLOAD"], display)
        return

    # Normal case: per-run replacement to preserve formatting
    for run in paragraph.runs:
        text = run.text
        for key, val in non_link_keys.items():
            text = text.replace(f'[{key}]', val or "")
        run.text = text

    # Check for remaining placeholders (spanning runs)
    new_full_text = "".join([r.text for r in paragraph.runs])
    remaining = PLACEHOLDER_PATTERN.findall(new_full_text)
    if remaining:
        # Fallback: rebuild with single run (loses formatting for spanning parts)
        new_text = new_full_text
        for key, val in non_link_keys.items():
            new_text = new_text.replace(f'[{key}]', val or "")
        # Clear runs
        for i in range(len(paragraph.runs) - 1, -1, -1):
            paragraph._element.remove(paragraph.runs[i]._element)
        # Add new run
        paragraph.add_run(new_text)
        print(f"Warning: Rebuilt paragraph due to spanning placeholders: {remaining}", file=sys.stderr)

def replace_in_table(table, mapping: Dict[str, str]):
    for row in table.rows:
        for cell in row.cells:
            replace_in_block(cell, mapping)

def replace_in_block(block, mapping: Dict[str, str]):
    for para in block.paragraphs:
        replace_in_paragraph(para, mapping)
    for table in getattr(block, "tables", []):
        replace_in_table(table, mapping)

def process_document(template_path: str, output_path: str, mapping: Dict[str, str]):
    doc = Document(template_path)
    replace_in_block(doc, mapping)
    for section in doc.sections:
        if section.header:
            replace_in_block(section.header, mapping)
        if section.footer:
            replace_in_block(section.footer, mapping)
    doc.save(output_path)

def fix_docx_templates():
    """Fix para PyInstaller não incluir templates docx"""
    import sys
    import os
    if hasattr(sys, '_MEIPASS'):
        # Estamos executando via PyInstaller
        import docx
        template_dir = os.path.join(sys._MEIPASS, 'docx', 'templates')
        if os.path.exists(template_dir):
            docx.shared.TEMPLATE_DIR = template_dir

# ----------------- CLI flow -----------------
def run_cli(template: Optional[str] = None, cnpj: Optional[str] = None, drive: Optional[str] = None,
            drive_text: Optional[str] = None, use_ai: Optional[bool] = None, ai_provider: Optional[str] = None,
            out: Optional[str] = None, extra_mapping: Optional[dict] = None) -> None:
    try:
        if not template:
            template = input("Caminho do template .docx: ").strip()
        if not Path(template).exists():
            print("Template não encontrado:", template)
            return
        if not cnpj:
            cnpj = input("CNPJ da empresa: ").strip()
        try:
            cnpj_norm = normalize_cnpj(cnpj)
        except Exception as e:
            print("CNPJ inválido:", e)
            return
        print("Consultando ReceitaWS...")
        data = consulta_empresa(cnpj_norm)
        mapping = build_mapping(data)
        if drive is None:
            drive = input("Link do Drive (opcional, ENTER para pular): ").strip()
        if drive:
            if not drive.startswith(("http://", "https://")):
                drive = "https://" + drive
            mapping["LINK_DRIVE"] = drive
            mapping["LINK_DRIVE_TEXT"] = drive_text or input("Texto do link (ENTER para 'Link Drive'): ").strip() or "Link Drive"
            mapping["LINK_PARA_DOWNLOAD"] = drive
            mapping["LINK_PARA_DOWNLOAD_TEXT"] = "Link para download"

        # Handle extra fields
        extra_fields = {
            "DATA_BACKUP": "Data Backup (opcional): ",
            "DATA_KICKOFF": "Data Kickoff (opcional): ",
            "DATA_ENTREGA": "Data Entrega (opcional): ",
            "DOMINIO": "Domínio (opcional): ",
            "DEMANDA": "Demanda (opcional): ",
        }
        for field, prompt in extra_fields.items():
            value = extra_mapping.get(field, "") if extra_mapping else ""
            if not value:
                value = input(prompt).strip()
            mapping[field] = value

        # Auto-set DOMINIOWP based on DOMINIO
        if mapping.get("DOMINIO"):
            mapping["DOMINIOWP"] = mapping["DOMINIO"] + "/wp-admin/"
        else:
            mapping["DOMINIOWP"] = ""

        # ESPECIALISTARESPONSAVEL is always fixed
        mapping["ESPECIALISTARESPONSAVEL"] = "ITALO GOMES"

        if use_ai is None:
            use_ai = input("Deseja usar IA para preencher [OBJETIVO_EMPRESA]? (s/N): ").strip().lower() == 's'
        if use_ai:
            provider = (ai_provider or os.environ.get("AI_PROVIDER") or input("Provedor IA (mock/hf/openai) [mock]: ").strip() or "mock")
            try:
                ai = get_ai_provider(provider)
            except Exception as e:
                print("Erro ao inicializar provedor IA:", e)
                print("Usando MockProvider como fallback.")
                ai = MockProvider()
            source_parts = []
            if mapping.get("ATIVIDADE_PRINCIPAL"):
                source_parts.append("Atividade principal: " + mapping["ATIVIDADE_PRINCIPAL"])
            if mapping.get("RESUMO_EMPRESA_CLIENTE"):
                source_parts.append("Resumo: " + mapping["RESUMO_EMPRESA_CLIENTE"])
            source_text = "\n".join(source_parts).strip() or str(data)[:2000]
            try:
                mapping["OBJETIVO_EMPRESA"] = ai.generate_objective(source_text, mapping)
            except Exception as e:
                print("Erro ao gerar objetivo com IA:", e)
                print("Usando heurística local.")
                mapping["OBJETIVO_EMPRESA"] = MockProvider().generate_objective(source_text, mapping)
        else:
            mapping["OBJETIVO_EMPRESA"] = ""
        out_path = out or input("Arquivo de saída (.docx) [relatorio_saida.docx]: ").strip() or f'relatorio_{cnpj_norm}.docx'
        print("Gerando documento...")
        process_document(template, out_path, mapping)
        print("Documento gerado:", out_path)
    except Exception as e:
        print("Erro durante execução:", e)

# ----------------- GUI flow -----------------
if TKINTER_AVAILABLE:
    class App:
        def __init__(self, root):
            self.root = root
            root.title("Gerador de Relatório - CNPJ -> Word")
            frm = tk.Frame(root, padx=10, pady=10)
            frm.pack(fill=tk.BOTH, expand=True)
            # Template
            tk.Label(frm, text="Template (.docx):").grid(row=0, column=0, sticky='w')
            self.entry_template = tk.Entry(frm, width=60)
            self.entry_template.grid(row=0, column=1, sticky='w')
            tk.Button(frm, text="Abrir", command=self.browse_template).grid(row=0, column=2)
            # CNPJ
            tk.Label(frm, text="CNPJ da empresa:").grid(row=1, column=0, sticky='w', pady=(10,0))
            self.entry_cnpj = tk.Entry(frm, width=40)
            self.entry_cnpj.grid(row=1, column=1, sticky='w', pady=(10,0))
            # Link Drive
            tk.Label(frm, text="Link do Drive (opcional):").grid(row=2, column=0, sticky='w')
            self.entry_drive = tk.Entry(frm, width=60)
            self.entry_drive.grid(row=2, column=1, sticky='w')
            tk.Label(frm, text="Texto do Link:").grid(row=2, column=2, sticky='w')
            self.entry_drive_text = tk.Entry(frm, width=20)
            self.entry_drive_text.grid(row=2, column=3, sticky='w')
            # Extra fields
            tk.Label(frm, text="Data Backup:").grid(row=3, column=0, sticky='w')
            self.entry_data_backup = tk.Entry(frm, width=20); self.entry_data_backup.grid(row=3, column=1, sticky='w')
            tk.Label(frm, text="Data Kickoff:").grid(row=4, column=0, sticky='w')
            self.entry_data_kickoff = tk.Entry(frm, width=20); self.entry_data_kickoff.grid(row=4, column=1, sticky='w')
            tk.Label(frm, text="Data Entrega:").grid(row=5, column=0, sticky='w')
            self.entry_data_entrega = tk.Entry(frm, width=20); self.entry_data_entrega.grid(row=5, column=1, sticky='w')
            tk.Label(frm, text="Domínio:").grid(row=6, column=0, sticky='w')
            self.entry_dominio = tk.Entry(frm, width=40); self.entry_dominio.grid(row=6, column=1, sticky='w')
            tk.Label(frm, text="Demanda:").grid(row=7, column=0, sticky='w')
            self.entry_demanda = tk.Entry(frm, width=40); self.entry_demanda.grid(row=7, column=1, sticky='w')
            # Removed DOMINIOWP entry - auto-generated
            tk.Label(frm, text="Especialista Responsável:").grid(row=8, column=0, sticky='w')
            self.entry_especialista = tk.Entry(frm, width=40); self.entry_especialista.grid(row=8, column=1, sticky='w')
            self.entry_especialista.insert(0, "ITALO GOMES")
            self.entry_especialista.config(state="readonly")  # Make it read-only
            # IA
            self.use_ai_var = tk.IntVar(value=1)
            tk.Checkbutton(frm, text="Usar IA para preencher [OBJETIVO_EMPRESA]", variable=self.use_ai_var).grid(row=9, column=0, sticky='w', columnspan=2, pady=(10,0))
            tk.Label(frm, text="Provedor IA:").grid(row=10, column=0, sticky='w', pady=(10,0))
            self.ai_provider = tk.StringVar(value=os.environ.get('AI_PROVIDER', 'mock'))
            tk.OptionMenu(frm, self.ai_provider, 'mock', 'hf', 'openai').grid(row=10, column=1, sticky='w')
            # Arquivo saída
            tk.Label(frm, text="Arquivo saída (.docx):").grid(row=11, column=0, sticky='w', pady=(10,0))
            self.entry_out = tk.Entry(frm, width=60)
            self.entry_out.grid(row=11, column=1, sticky='w')
            self.entry_out.insert(0, 'relatorio_saida.docx')
            # Botão gerar
            tk.Button(frm, text="Gerar Relatório", command=self.run).grid(row=12, column=1, pady=20)

        def browse_template(self):
            p = filedialog.askopenfilename(filetypes=[('Word files', '*.docx')])
            if p:
                self.entry_template.delete(0, tk.END)
                self.entry_template.insert(0, p)

        def run(self):
            template = self.entry_template.get().strip()
            if not template or not Path(template).exists():
                messagebox.showerror('Erro', 'Template .docx inválido ou não informado')
                return
            cnpj_raw = self.entry_cnpj.get().strip()
            try:
                cnpj_norm = normalize_cnpj(cnpj_raw)
            except Exception as e:
                messagebox.showerror('Erro', f'CNPJ inválido: {e}')
                return
            try:
                data = consulta_empresa(cnpj_norm)
            except Exception as e:
                messagebox.showerror('Erro', f'Falha ao consultar ReceitaWS: {e}')
                return
            mapping = build_mapping(data)
            mapping["DATA_BACKUP"] = self.entry_data_backup.get().strip()
            mapping["DATA_KICKOFF"] = self.entry_data_kickoff.get().strip()
            mapping["DATA_ENTREGA"] = self.entry_data_entrega.get().strip()
            dominio = self.entry_dominio.get().strip()
            mapping["DOMINIO"] = dominio
            if dominio:
                mapping["DOMINIOWP"] = dominio + "/wp-admin/"
            else:
                mapping["DOMINIOWP"] = ""
            mapping["DEMANDA"] = self.entry_demanda.get().strip()
            mapping["ESPECIALISTARESPONSAVEL"] = "ITALO GOMES"
            drive = self.entry_drive.get().strip()
            if drive:
                if not drive.startswith(('http://', 'https://')):
                    drive = 'https://' + drive
                mapping['LINK_DRIVE'] = drive
                mapping['LINK_DRIVE_TEXT'] = self.entry_drive_text.get().strip() or 'Link Drive'
                mapping['LINK_PARA_DOWNLOAD'] = drive
                mapping['LINK_PARA_DOWNLOAD_TEXT'] = 'Link para download'
            if self.use_ai_var.get():
                source_parts = []
                if mapping.get('ATIVIDADE_PRINCIPAL'):
                    source_parts.append('Atividade principal: ' + mapping['ATIVIDADE_PRINCIPAL'])
                if mapping.get('RESUMO_EMPRESA_CLIENTE'):
                    source_parts.append('Resumo: ' + mapping['RESUMO_EMPRESA_CLIENTE'])
                source_text = '\n'.join(source_parts).strip() or str(data)[:2000]
                try:
                    ai = get_ai_provider(self.ai_provider.get())
                except Exception as e:
                    messagebox.showwarning('Aviso', f'Erro ao inicializar provedor IA: {e}\nUsando mock.')
                    ai = MockProvider()
                try:
                    objetivo = ai.generate_objective(source_text, mapping)
                except Exception as e:
                    messagebox.showwarning('Aviso', f'Erro ao gerar objetivo com IA: {e}\nUsando heurística local.')
                    objetivo = MockProvider().generate_objective(source_text, mapping)
                mapping['OBJETIVO_EMPRESA'] = objetivo
            else:
                mapping['OBJETIVO_EMPRESA'] = ''
            out = self.entry_out.get().strip() or f'relatorio_{cnpj_norm}.docx'
            try:
                process_document(template, out, mapping)
            except Exception as e:
                messagebox.showerror('Erro', f'Erro ao processar documento: {e}')
                return
            messagebox.showinfo('Sucesso', f'Relatório gerado: {out}')

# ----------------- Main -----------------
def main():
    fix_docx_templates()
    parser = argparse.ArgumentParser(description="Preencher relatórios Word via CNPJ")
    parser.add_argument("--template", help=".docx template")
    parser.add_argument("--cnpj", help="CNPJ da empresa")
    parser.add_argument("--drive", help="Link do Drive")
    parser.add_argument("--drive-text", help="Texto do link do Drive")
    parser.add_argument("--use-ai", action="store_true", help="Usar IA para preencher [OBJETIVO_EMPRESA]")
    parser.add_argument("--ai-provider", help="Provedor IA: mock/hf/openai")
    parser.add_argument("--out", help="Arquivo de saída .docx")
    parser.add_argument("--data-backup", help="Data de backup [DATA_BACKUP]")
    parser.add_argument("--data-kickoff", help="Data de kickoff [DATA_KICKOFF]")
    parser.add_argument("--data-entrega", help="Data de entrega [DATA_ENTREGA]")
    parser.add_argument("--dominio", help="Texto [DOMINIO]")
    parser.add_argument("--demanda", help="Texto [DEMANDA]")
    parser.add_argument("--run-tests", action="store_true", help="Executar testes rápidos")
    args = parser.parse_args()
    if args.run_tests:
        print("Testes simples:")
        try:
            assert normalize_cnpj("12.345.678/0001-95") == "12345678000195"
            assert normalize_cnpj("12345678000195") == "12345678000195"
            print("normalize_cnpj OK")
        except AssertionError:
            print("normalize_cnpj falhou")
        try:
            doc = Document()
            p = doc.add_paragraph("[NOME_EMPRESA_CLIENTE] e [CNPJ]")
            mapping = {"NOME_EMPRESA_CLIENTE":"ACME","CNPJ":"123"}
            replace_in_paragraph(p, mapping)
            assert "ACME" in p.text and "123" in p.text
            print("replace_in_paragraph OK")
        except Exception as e:
            print("replace_in_paragraph falhou:", e)
        return
    if TKINTER_AVAILABLE and not any([args.template, args.cnpj, args.drive]):
        root = tk.Tk()
        app = App(root)
        root.mainloop()
    else:
        extra_mapping = {
            "DATA_BACKUP": args.data_backup or "",
            "DATA_KICKOFF": args.data_kickoff or "",
            "DATA_ENTREGA": args.data_entrega or "",
            "DOMINIO": args.dominio or "",
            "DEMANDA": args.demanda or "",
        }
        run_cli(
            template=args.template,
            cnpj=args.cnpj,
            drive=args.drive,
            drive_text=args.drive_text,
            use_ai=args.use_ai,
            ai_provider=args.ai_provider,
            out=args.out,
            extra_mapping=extra_mapping
        )

if __name__ == "__main__":
    main()