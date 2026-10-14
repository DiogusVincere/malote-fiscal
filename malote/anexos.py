"""De qual empresa é a nota? Procura o CNPJ das empresas candidatas no texto do anexo.

Só é usado quando o mesmo e-mail serve a várias empresas (o contador de um grupo, o dono de
duas lojas). Na NFS-e aparecem o CNPJ do prestador e o do tomador; o prestador não é cliente, então
basta ver qual dos candidatos aparece. Mais de um candidato no mesmo arquivo não decide nada.
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

_CNPJ = re.compile(r"(?<!\d)(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})(?!\d)")
LIMITE = 15 * 1024 * 1024  # PDF maior que isso não vale o tempo


def _texto_pdf(dados: bytes) -> str:
    try:
        from pdfminer.high_level import extract_text
        return extract_text(io.BytesIO(dados), maxpages=3) or ""
    except Exception:
        return ""


def _texto(nome: str, dados: bytes) -> str:
    sufixo = Path(nome).suffix.lower()
    if len(dados) > LIMITE:
        return ""
    if sufixo == ".pdf":
        return _texto_pdf(dados)
    if sufixo in (".xml", ".txt", ".html", ".htm"):
        return dados.decode("utf-8", errors="ignore")
    if sufixo == ".zip":
        partes = []
        try:
            with zipfile.ZipFile(io.BytesIO(dados)) as z:
                for info in z.infolist()[:60]:
                    if Path(info.filename).suffix.lower() in (".pdf", ".xml") and info.file_size <= LIMITE:
                        partes.append(_texto(info.filename, z.read(info)))
        except (zipfile.BadZipFile, OSError):
            pass
        return "\n".join(partes)
    return ""


def cnpjs(caminho: Path, nome: str) -> set[str]:
    return {re.sub(r"\D", "", m) for m in _CNPJ.findall(_texto(nome, caminho.read_bytes()))}


def qual_empresa(achados: set[str], candidatos: dict[str, list[str]]) -> str | None:
    """candidatos: {pasta: [cnpjs]}. Devolve a pasta cujo CNPJ aparece, se for uma só."""
    batem = [pasta for pasta, lista in candidatos.items() if any(c and c in achados for c in lista)]
    if len(batem) == 1:
        return batem[0]
    if not batem:  # filial com CNPJ diferente: tenta pela raiz (8 primeiros dígitos)
        raizes = {a[:8] for a in achados}
        batem = [pasta for pasta, lista in candidatos.items() if any(c and c[:8] in raizes for c in lista)]
        if len(batem) == 1:
            return batem[0]
    return None
