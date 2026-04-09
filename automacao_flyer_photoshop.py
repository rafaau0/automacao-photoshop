import sys
from pathlib import Path
import pandas as pd

# Requerimentos:
#   pip install pandas openpyxl pywin32
# Requer também:
#   Adobe Photoshop instalado no Windows
#
# Estrutura esperada do projeto:
# AUTOMATIZA PHOTOSHOP/
# ├─ CSV/
# │  └─ dados_flyer.xlsx
# ├─ PSD/
# │  └─ 9 ITENS AÇOUGUE copiar.psd
# ├─ automacao_flyer_photoshop.py
# └─ saida/

try:
    import win32com.client
except ImportError:
    print("Erro: pywin32 não está instalado.")
    print("Instale com: pip install pywin32")
    sys.exit(1)


BASE_DIR = Path(__file__).resolve().parent
PSD_PATH = BASE_DIR / "PSD" / "9 ITENS AÇOUGUE.psd"
PLANILHA_PATH = BASE_DIR / "CSV" / "dados_flyer.xlsx"
PASTA_SAIDA = BASE_DIR / "saida"

NOME_BASE_EXPORTACAO = "flyer_atualizado"
ABA_PLANILHA = "Página1"
EXPORTAR_JPG = True
QUALIDADE_JPG = 10
SALVAR_COPIA_PSD = True

# Nome do grupo principal no PSD
NOME_GRUPO = "09 ITENS copiar 3"

# Prefixos das camadas no PSD
PREFIXO_DESCRICAO = "desc-produto-"
PREFIXO_PRECO = "preco-produto-"

# Se True, também tenta atualizar preço
ATUALIZAR_PRECO = True


def normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def formatar_indice(numero: int) -> str:
    return f"{numero:02d}"


def carregar_mapa_textos(arquivo_excel: Path, aba="Página1") -> dict:
    df = pd.read_excel(arquivo_excel, sheet_name=aba)
    df.columns = [str(c).strip() for c in df.columns]

    colunas_norm = {str(c).lower().strip(): c for c in df.columns}

    if "descrição" not in colunas_norm and "descricao" not in colunas_norm:
        print("\nColunas encontradas na planilha:")
        for c in df.columns:
            print(f"- {c}")
        raise ValueError("A planilha precisa ter a coluna 'descrição' ou 'descricao'.")

    col_desc = colunas_norm.get("descrição") or colunas_norm.get("descricao")
    col_preco = colunas_norm.get("preco") or colunas_norm.get("preço")

    mapa = {}

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        indice = formatar_indice(i)

        descricao = normalizar_texto(row[col_desc])
        if descricao:
            # remove preço duplicado no final da descrição, se existir
            if col_preco:
                preco = normalizar_texto(row[col_preco])
                if preco and descricao.endswith(preco):
                    descricao = descricao[: -len(preco)].strip()

            mapa[f"{PREFIXO_DESCRICAO}{indice}"] = descricao

        if ATUALIZAR_PRECO and col_preco:
            preco = normalizar_texto(row[col_preco])
            if preco:
                mapa[f"{PREFIXO_PRECO}{indice}"] = preco

    if not mapa:
        raise ValueError("Nenhum dado válido foi encontrado na planilha.")

    return mapa


def conectar_photoshop():
    app = win32com.client.Dispatch("Photoshop.Application")
    app.DisplayDialogs = 3
    return app


def iterar_camadas(container):
    try:
        for i in range(1, container.ArtLayers.Count + 1):
            yield container.ArtLayers.Item(i)
    except Exception:
        pass

    try:
        for i in range(1, container.LayerSets.Count + 1):
            grupo = container.LayerSets.Item(i)
            yield grupo
            for item in iterar_camadas(grupo):
                yield item
    except Exception:
        pass


def buscar_grupo_por_nome(doc, nome_grupo: str):
    if not nome_grupo:
        return doc

    for item in iterar_camadas(doc):
        try:
            if hasattr(item, "Name") and item.Name == nome_grupo:
                return item
        except Exception:
            continue

    raise ValueError(f"Grupo '{nome_grupo}' não encontrado no PSD.")


def buscar_camada_texto_por_nome(container, nome_camada: str):
    for item in iterar_camadas(container):
        try:
            if hasattr(item, "Name") and item.Name == nome_camada:
                _ = item.TextItem
                return item
        except Exception:
            continue
    return None


def atualizar_textos(doc, mapa_textos: dict, nome_grupo: str | None = None):
    area_busca = buscar_grupo_por_nome(doc, nome_grupo)
    nao_encontradas = []

    for nome_camada, novo_texto in mapa_textos.items():
        layer = buscar_camada_texto_por_nome(area_busca, nome_camada)

        if layer is None:
            nao_encontradas.append(nome_camada)
            continue

        try:
            layer.TextItem.contents = novo_texto
            print(f"OK -> {nome_camada}: {novo_texto}")
        except Exception as e:
            print(f"ERRO -> {nome_camada}: {e}")

    return nao_encontradas


def salvar_psd_como(doc, caminho_saida: Path):
    psd_options = win32com.client.Dispatch("Photoshop.PhotoshopSaveOptions")
    doc.SaveAs(str(caminho_saida), psd_options, True)


def exportar_jpg(doc, caminho_saida_jpg: Path, qualidade: int = 10):
    jpg_options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
    jpg_options.Quality = qualidade
    jpg_options.EmbedColorProfile = True
    jpg_options.Matte = 1
    doc.SaveAs(str(caminho_saida_jpg), jpg_options, True)


def validar_arquivos():
    if not PSD_PATH.exists():
        raise FileNotFoundError(
            f"PSD não encontrado em: {PSD_PATH}\n"
            f"Coloque o arquivo dentro da pasta: {BASE_DIR / 'PSD'}"
        )

    if not PLANILHA_PATH.exists():
        raise FileNotFoundError(
            f"Planilha não encontrada em: {PLANILHA_PATH}\n"
            f"Coloque o arquivo dentro da pasta: {BASE_DIR / 'CSV'}"
        )


def main():
    print("Iniciando automação do flyer...")
    print(f"Base do projeto: {BASE_DIR}")

    validar_arquivos()
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    mapa_textos = carregar_mapa_textos(PLANILHA_PATH, ABA_PLANILHA)

    print("\nTextos carregados da planilha:")
    for camada, texto in mapa_textos.items():
        print(f"- {camada}: {texto}")

    app = conectar_photoshop()
    doc = app.Open(str(PSD_PATH))

    try:
        nao_encontradas = atualizar_textos(doc, mapa_textos, NOME_GRUPO)

        psd_saida = PASTA_SAIDA / f"{NOME_BASE_EXPORTACAO}.psd"
        jpg_saida = PASTA_SAIDA / f"{NOME_BASE_EXPORTACAO}.jpg"

        if SALVAR_COPIA_PSD:
            salvar_psd_como(doc, psd_saida)
            print(f"\nPSD salvo em: {psd_saida}")

        if EXPORTAR_JPG:
            exportar_jpg(doc, jpg_saida, QUALIDADE_JPG)
            print(f"JPG salvo em: {jpg_saida}")

        if nao_encontradas:
            print("\nAtenção: algumas camadas não foram encontradas:")
            for nome in nao_encontradas:
                print(f"- {nome}")
        else:
            print("\nTodas as camadas foram atualizadas com sucesso.")

    finally:
        # Fecha sem salvar por cima do original
        doc.Close(2)

    print("\nProcesso finalizado.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nErro na execução: {e}")
        sys.exit(1)