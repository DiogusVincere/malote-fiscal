r"""Monta o pacote portátil para Windows: a pasta "Malote Fiscal" e o zip dela.

    python build\empacotar.py
        -> dist\Malote Fiscal\  e  dist\malote-fiscal-v1.0.0-windows.zip
    python build\empacotar.py --versao 1.0.1 --sem-zip

Precisa do Python com as bibliotecas do requirements.txt, do PyInstaller e do Pillow (só para o ícone):
    pip install pyinstaller pillow
"""
import argparse
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

RAIZ = Path(os.path.abspath(__file__)).parent.parent
NOME = "Malote Fiscal"

LEIA_ME = r"""MALOTE FISCAL - serviços tomados por e-mail (versão {versao})

O que é
Uma tela que roda ao lado do Outlook desta máquina: pergunta aos clientes do escritório se tomaram
serviços no mês, lê TODAS as respostas (inclusive as que não parecem resposta), guarda as notas na
pasta de cada cliente e ajuda a responder. Nada de servidor, senha ou IA de fora.

Como usar
1. Descompacte este zip numa pasta qualquer.
2. Abra "Malote Fiscal.exe". Na primeira vez o Windows pode avisar que o editor é desconhecido
   (SmartScreen): clique em "Mais informações" e depois em "Executar assim mesmo".
3. Escolha a conta do Outlook no alto da tela, a pasta onde cada cliente tem a sua subpasta
   ("<código> - <nome>") e a planilha de clientes (Código, Razão Social, CNPJ/CPF, E-mails,
   Regime municipal, Resp. fiscal, Ex-cliente), normalmente em Downloads.

Sem Outlook à mão
"Malote Fiscal (demonstracao).bat" (ou a opção "demonstração" na lista de contas) abre a mesma
tela sobre um Outlook simulado, com um escritório e clientes fictícios. Nada sai da máquina.

Onde ficam as coisas
- Configuração, cadastro de contatos, respostas aprendidas e as planilhas de controle:
  %LOCALAPPDATA%\Malote Fiscal. Pode apagar essa pasta quando quiser.
- As notas vão para a pasta de cada cliente: <cliente>\NFE\<ano>\<mês>\Tomados\Tomados e-mail.
  Nada é sobrescrito.
- A demonstração monta tudo em %TEMP%\malote-fiscal-demo e recomeça a cada abertura.

Código-fonte e documentação: https://github.com/DiogusVincere/malote-fiscal
Licença MIT (LICENCA.txt).
"""

BAT_DEMO = '@echo off\r\nstart "" "%~dp0Malote Fiscal.exe" --demo\r\n'


def rodar(comando, **kw):
    print(">", " ".join(str(c) for c in comando))
    subprocess.run([str(c) for c in comando], check=True, **kw)


def congelar(saida: Path, trabalho: Path, versao: str) -> Path:
    ambiente = dict(os.environ, MALOTE_VERSAO=versao)
    rodar([sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
           "--distpath", saida, "--workpath", trabalho, RAIZ / "build" / "malote.spec"], env=ambiente, cwd=str(RAIZ))
    pacote = saida / NOME
    if not (pacote / f"{NOME}.exe").is_file():
        raise SystemExit(f"O PyInstaller não gerou {pacote / (NOME + '.exe')}.")
    return pacote


def documentos(pacote: Path, versao: str) -> None:
    (pacote / "LEIA-ME.txt").write_text(LEIA_ME.format(versao=versao), encoding="utf-8-sig", newline="\r\n")
    (pacote / "Malote Fiscal (demonstracao).bat").write_text(BAT_DEMO, encoding="cp1252", newline="")
    shutil.copy(RAIZ / "LICENSE", pacote / "LICENCA.txt")


def compactar(pacote: Path, zip_destino: Path) -> Path:
    if zip_destino.exists():
        zip_destino.unlink()
    with zipfile.ZipFile(zip_destino, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for arquivo in sorted(pacote.rglob("*")):
            if arquivo.is_file():
                zf.write(arquivo, Path(NOME) / arquivo.relative_to(pacote))
    return zip_destino


def main() -> None:
    p = argparse.ArgumentParser(description="Monta o pacote portátil do Malote Fiscal.")
    p.add_argument("--versao", default="1.0.0")
    p.add_argument("--saida", default=str(RAIZ / "dist"))
    p.add_argument("--trabalho", default=str(RAIZ / "build" / "pyinstaller"))
    p.add_argument("--sem-zip", action="store_true")
    args = p.parse_args()

    saida, trabalho = Path(args.saida), Path(args.trabalho)
    saida.mkdir(parents=True, exist_ok=True)
    pacote = congelar(saida, trabalho, args.versao)
    documentos(pacote, args.versao)
    tamanho = sum(f.stat().st_size for f in pacote.rglob("*") if f.is_file())
    print(f"\nPacote: {pacote}  ({tamanho / 1e6:.1f} MB)")
    if not args.sem_zip:
        destino = compactar(pacote, saida / f"malote-fiscal-v{args.versao}-windows.zip")
        print(f"Zip: {destino}  ({destino.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
