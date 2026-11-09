# -*- coding: utf-8 -*-
"""
Gerador de documentos (CASSOL)
  - Termo de Comodato - Notebook  (modelo_termo.docx)
  - PDF de Boas-Vindas            (modelo_boas_vindas.pptx)

Conversão para PDF (nesta ordem):
  1) Microsoft Word / PowerPoint instalado (Windows) -> resultado idêntico ao original
  2) LibreOffice (soffice) -> alternativa gratuita
"""
import os, re, sys, shutil, zipfile, tempfile, subprocess, datetime
from xml.sax.saxutils import escape

PASTA = os.path.dirname(os.path.abspath(sys.argv[0] if getattr(sys, "frozen", False) else __file__))
MODELO = os.path.join(PASTA, "modelo_termo.docx")
MODELO_BV = os.path.join(PASTA, "modelo_boas_vindas.pptx")
DOMINIO = "cassol.ind.br"
SAIDA = os.path.join(PASTA, "PDFs gerados")

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
         "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# Regex de um campo de formulário: do "separate" até o "end" (o texto entre eles é o valor)
_CAMPO = re.compile(
    r'(fldCharType="separate"/></w:r>)(.*?)'
    r'(<w:r(?: [^>]*)?>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<w:fldChar w:fldCharType="end"/>)',
    re.S)


_RUN_TEXTO = re.compile(r"<w:r(?: [^>]*)?>(?:(?!</w:r>).)*?<w:t[ >](?:(?!</w:r>).)*</w:r>", re.S)


def preencher_docx(destino, d):
    """d: dict com dia, mes, ano2, nome, documento, modelo, serie, patrimonio"""
    # ordem dos campos no documento
    valores = [d["dia"], d["mes"], d["ano2"], d["nome"], d["documento"],
               d["modelo"], d["serie"], d["patrimonio"], d["nome"]]
    with zipfile.ZipFile(MODELO) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        itens = [(i, z.read(i.filename)) for i in z.infolist()]

    it = iter(valores)

    def troca(m):
        try:
            v = next(it)
        except StopIteration:
            return m.group(0)
        meio = m.group(2)
        rpr = re.search(r"<w:rPr>.*?</w:rPr>", meio, re.S)
        rpr = rpr.group(0) if rpr else ""
        run_novo = f'<w:r>{rpr}<w:t xml:space="preserve">{escape(v)}</w:t></w:r>'
        primeiro = [True]

        def so_texto(r):  # troca o 1º run com texto; remove os demais; mantém quebras de parágrafo
            if primeiro[0]:
                primeiro[0] = False
                return run_novo
            return ""
        meio = _RUN_TEXTO.sub(so_texto, meio)
        if primeiro[0]:  # campo vazio no modelo
            meio = run_novo + meio
        return m.group(1) + meio + m.group(3)

    novo, n = _CAMPO.subn(troca, xml)
    if n != len(valores):
        raise RuntimeError(f"O modelo tem {n} campos, esperado {len(valores)}. "
                           "O modelo_termo.docx foi alterado?")
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in itens:
            if info.filename == "word/document.xml":
                data = novo.encode("utf-8")
            out.writestr(info, data)


def _achar_soffice():
    for nome in ("soffice", "libreoffice"):
        p = shutil.which(nome)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/usr/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def _pdf_via_com(src, pdf, app):
    import win32com.client  # pip install pywin32
    import pythoncom
    pythoncom.CoInitialize()
    src, pdf = os.path.abspath(src), os.path.abspath(pdf)
    if app == "Word":
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Open(src, ReadOnly=True)
            doc.ExportAsFixedFormat(pdf, 17)  # 17 = PDF
            doc.Close(False)
        finally:
            word.Quit()
    else:
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            pres = ppt.Presentations.Open(src, True, False, False)  # somente leitura, sem janela
            pres.SaveAs(pdf, 32)  # 32 = PDF
            pres.Close()
        finally:
            ppt.Quit()


def _pdf_via_libreoffice(src, pdf):
    exe = _achar_soffice()
    if not exe:
        raise RuntimeError("LibreOffice não encontrado.")
    pasta = tempfile.mkdtemp()
    try:
        subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", pasta, src],
                       check=True, capture_output=True, timeout=180)
        gerado = os.path.join(pasta, os.path.splitext(os.path.basename(src))[0] + ".pdf")
        shutil.move(gerado, pdf)
    finally:
        shutil.rmtree(pasta, ignore_errors=True)


def _pdf_via_powershell(src, pdf, app):
    """Usa o Word/PowerPoint pelo PowerShell (já vem no Windows) - dispensa o pywin32."""
    if app == "Word":
        ps = ("$a = New-Object -ComObject Word.Application; $a.Visible = $false; $a.DisplayAlerts = 0; "
              "try { $d = $a.Documents.Open($env:DOC_IN); "
              "$d.ExportAsFixedFormat($env:DOC_OUT, 17); $d.Close($false) } finally { $a.Quit() }")
    else:
        ps = ("$a = New-Object -ComObject PowerPoint.Application; "
              "try { $d = $a.Presentations.Open($env:DOC_IN, -1, 0, 0); "
              "$d.SaveAs($env:DOC_OUT, 32); $d.Close() } finally { $a.Quit() }")
    env = dict(os.environ, DOC_IN=os.path.abspath(src), DOC_OUT=os.path.abspath(pdf))
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                       env=env, capture_output=True, text=True, timeout=180,
                       creationflags=0x08000000)  # sem janela preta
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout or "falha desconhecida").strip()[:400])


def _converter(src, pdf, app="Word"):
    erros = []
    if sys.platform.startswith("win"):
        for nome, fn in ((f"{app} (pywin32)", lambda a, b: _pdf_via_com(a, b, app)),
                         (f"{app} (PowerShell)", lambda a, b: _pdf_via_powershell(a, b, app))):
            try:
                fn(src, pdf)
                if os.path.exists(pdf):
                    return
            except Exception as ex:
                erros.append(f"{nome}: {ex}")
    try:
        _pdf_via_libreoffice(src, pdf)
        if os.path.exists(pdf):
            return
    except Exception as ex:
        erros.append(f"LibreOffice: {ex}")
    raise RuntimeError("Não foi possível converter para PDF.\n\n"
                       f"Este computador precisa ter o Microsoft {app} (versão instalada) "
                       "ou o LibreOffice.\n\nDetalhes técnicos:\n" + "\n".join(erros))


# ------------------------------------------------------------ Boas-Vindas (PPTX)
def primeiro_nome(texto):
    t = texto.split(",")[0].strip()
    return (t.split()[0] if t else "").capitalize()


def normalizar_usuario(texto):
    return texto.strip().split("@")[0].strip().lower()


def preencher_pptx(destino, nome, usuario):
    with zipfile.ZipFile(MODELO_BV) as z:
        itens = [(i, z.read(i.filename)) for i in z.infolist()]
    trocas = [("<a:t>Vitor,</a:t>", f"<a:t>{escape(nome)},</a:t>"),
              ("<a:t>vpqueiroz</a:t>", f"<a:t>{escape(usuario)}</a:t>"),
              ("<a:t> vpqueiroz@</a:t>", f"<a:t> {escape(usuario)}@</a:t>")]
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in itens:
            if info.filename == "ppt/slides/slide1.xml":
                xml = data.decode("utf-8")
                for velho, novo in trocas:
                    if velho not in xml:
                        raise RuntimeError("O modelo_boas_vindas.pptx foi alterado: não achei " + velho)
                    xml = xml.replace(velho, novo)
                data = xml.encode("utf-8")
            out.writestr(info, data)


def gerar_boas_vindas(nome_digitado, usuario_digitado):
    nome, usuario = primeiro_nome(nome_digitado), normalizar_usuario(usuario_digitado)
    if not nome or not usuario:
        raise RuntimeError("Informe o nome e o usuário.")
    os.makedirs(SAIDA, exist_ok=True)
    seguro = re.sub(r'[\\/:*?"<>|]', "", usuario)
    pdf = os.path.join(SAIDA, f"Boas_vindas_{seguro}.pdf")
    tmp = os.path.join(tempfile.mkdtemp(), "boas_vindas.pptx")
    preencher_pptx(tmp, nome, usuario)
    try:
        _converter(tmp, pdf, "PowerPoint")
    finally:
        shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)
    return pdf


def gerar_pdf(d):
    os.makedirs(SAIDA, exist_ok=True)
    seguro = re.sub(r'[\\/:*?"<>|]', "", d["nome"]).strip().replace(" ", "_")
    pdf = os.path.join(SAIDA, f"Termo_de_comodato_notebook_{seguro}.pdf")
    tmp = os.path.join(tempfile.mkdtemp(), "termo.docx")
    preencher_docx(tmp, d)
    try:
        _converter(tmp, pdf, "Word")
    finally:
        shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)
    return pdf


def abrir(caminho):
    if sys.platform.startswith("win"):
        os.startfile(caminho)
    elif sys.platform == "darwin":
        subprocess.run(["open", caminho])
    else:
        subprocess.run(["xdg-open", caminho])


# ----------------------------------------------------------------- Interface
def interface():
    import tkinter as tk
    from tkinter import ttk, messagebox

    hoje = datetime.date.today()
    root = tk.Tk()
    root.title("Gerador de Documentos | CASSOL")
    root.resizable(False, False)
    try:
        ico = os.path.join(PASTA, "icone.ico")
        if sys.platform.startswith("win") and os.path.exists(ico):
            root.iconbitmap(ico)
        elif os.path.exists(os.path.join(PASTA, "icone.png")):
            root.iconphoto(True, tk.PhotoImage(file=os.path.join(PASTA, "icone.png")))
    except Exception:
        pass

    nb = ttk.Notebook(root)
    nb.pack(padx=10, pady=10)
    abrir_ao_final = tk.BooleanVar(value=True)

    def rodape(frm, row, ao_gerar, ao_limpar):
        ttk.Checkbutton(frm, text="Abrir o PDF ao terminar", variable=abrir_ao_final).grid(
            row=row, column=1, columnspan=3, sticky="w", pady=(8, 0))
        status = ttk.Label(frm, text="", foreground="#0a6", wraplength=420)
        status.grid(row=row + 2, column=0, columnspan=4, sticky="w", pady=(8, 0))
        bt = ttk.Frame(frm)
        bt.grid(row=row + 1, column=0, columnspan=4, pady=(12, 0), sticky="e")
        ttk.Button(bt, text="Limpar", command=ao_limpar).pack(side="left", padx=4)
        ttk.Button(bt, text="Gerar PDF", command=ao_gerar).pack(side="left", padx=4)
        ttk.Button(bt, text="Abrir pasta dos PDFs",
                   command=lambda: (os.makedirs(SAIDA, exist_ok=True), abrir(SAIDA))).pack(side="left", padx=4)
        return status

    def executar(status, funcao):
        status.config(text="Gerando PDF...", foreground="#555")
        root.update()
        try:
            pdf = funcao()
        except Exception as ex:
            status.config(text="")
            messagebox.showerror("Erro ao gerar", str(ex))
            return
        status.config(text="PDF gerado: " + os.path.basename(pdf), foreground="#0a6")
        if abrir_ao_final.get():
            abrir(pdf)

    # ---------------------------------------------------------- aba 1: Termo
    t = ttk.Frame(nb, padding=16)
    nb.add(t, text="Termo de Comodato")
    v = {k: tk.StringVar() for k in ("nome", "documento", "modelo", "serie", "patrimonio")}
    v["dia"] = tk.StringVar(value=f"{hoje.day:02d}")
    v["mes"] = tk.StringVar(value=MESES[hoje.month - 1])
    v["ano"] = tk.StringVar(value=str(hoje.year))
    v["modelo"].set("Notebook Dell Inspiron 3501")

    def linha(frm, r, rotulo, var, dica=None):
        ttk.Label(frm, text=rotulo).grid(row=r, column=0, sticky="w", pady=4)
        e = ttk.Entry(frm, textvariable=var, width=42)
        e.grid(row=r, column=1, columnspan=3, sticky="we", pady=4)
        if dica:
            ttk.Label(frm, text=dica, foreground="#777").grid(row=r + 1, column=1, columnspan=3, sticky="w")
        return e

    ttk.Label(t, text="Data").grid(row=1, column=0, sticky="w", pady=4)
    fd = ttk.Frame(t)
    fd.grid(row=1, column=1, columnspan=3, sticky="w")
    ttk.Entry(fd, textvariable=v["dia"], width=4).pack(side="left")
    ttk.Combobox(fd, textvariable=v["mes"], values=MESES, width=11, state="readonly").pack(side="left", padx=6)
    ttk.Entry(fd, textvariable=v["ano"], width=6).pack(side="left")
    e_nome = linha(t, 2, "Nome do colaborador", v["nome"])
    linha(t, 3, "Matrícula / CPF", v["documento"], "Exatamente como deve sair no termo. Ex.: CPF 000.000.000-00")
    linha(t, 5, "Modelo do equipamento", v["modelo"])
    linha(t, 6, "Número de série", v["serie"])
    linha(t, 7, "Número de patrimônio", v["patrimonio"])

    def limpar_termo():
        for k in ("nome", "documento", "serie", "patrimonio"):
            v[k].set("")
        e_nome.focus()

    def gerar_termo():
        d = {k: v[k].get().strip() for k in v}
        faltam = [r for k, r in (("nome", "Nome"), ("documento", "Matrícula/CPF"), ("modelo", "Modelo"),
                                 ("serie", "Número de série"), ("patrimonio", "Patrimônio"),
                                 ("dia", "Dia"), ("ano", "Ano")) if not d[k]]
        if faltam:
            messagebox.showwarning("Campos em branco", "Preencha: " + ", ".join(faltam))
            return
        d["ano2"] = d["ano"][-2:]
        d["dia"] = d["dia"].zfill(2)
        executar(st_termo, lambda: gerar_pdf(d))

    st_termo = rodape(t, 8, gerar_termo, limpar_termo)

    # ---------------------------------------------------- aba 2: Boas-Vindas
    b = ttk.Frame(nb, padding=16)
    nb.add(b, text="Boas-Vindas")
    bv = {"nome": tk.StringVar(), "usuario": tk.StringVar()}
    ttk.Label(b, text="Dados do novo colaborador", font=("Segoe UI", 11, "bold")).grid(
        row=0, columnspan=4, sticky="w", pady=(0, 6))
    e_bv = linha(b, 1, "Nome", bv["nome"], "Pode digitar o nome completo: só o primeiro nome sai no PDF")
    linha(b, 3, "Usuário", bv["usuario"], "Ex.: vpqueiroz (o e-mail usa o mesmo usuário)")
    previa = ttk.Label(b, text="", foreground="#444", justify="left")
    previa.grid(row=5, column=0, columnspan=4, sticky="w", pady=(10, 0))

    def atualizar(*_):
        n, u = primeiro_nome(bv["nome"].get()), normalizar_usuario(bv["usuario"].get())
        previa.config(text=f"Saudação: {n or '...'}, seja bem vindo(a)!\n"
                           f"Login de máquina: {u or '...'}\n"
                           f"E-mail: {(u or '...')}@{DOMINIO}")
    bv["nome"].trace_add("write", atualizar)
    bv["usuario"].trace_add("write", atualizar)
    atualizar()

    def limpar_bv():
        bv["nome"].set("")
        bv["usuario"].set("")
        e_bv.focus()

    def gerar_bv():
        if not primeiro_nome(bv["nome"].get()) or not normalizar_usuario(bv["usuario"].get()):
            messagebox.showwarning("Campos em branco", "Preencha o nome e o usuário.")
            return
        executar(st_bv, lambda: gerar_boas_vindas(bv["nome"].get(), bv["usuario"].get()))

    st_bv = rodape(b, 6, gerar_bv, limpar_bv)

    def enter(_e):
        (gerar_termo if nb.index(nb.select()) == 0 else gerar_bv)()
    root.bind("<Return>", enter)
    e_nome.focus()
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--teste":
        print(gerar_pdf(dict(dia="20", mes="Outubro", ano2="26", nome="Maria Souza Teste",
                             documento="CPF 111.222.333-44", modelo="Notebook Dell Latitude 5420",
                             serie="ABC1234", patrimonio="8001")))
    elif len(sys.argv) > 1 and sys.argv[1] == "--teste-bv":
        print(gerar_boas_vindas("maria Souza Teste", "msteste@cassol.ind.br"))
    else:
        interface()
