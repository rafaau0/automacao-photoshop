import os
from pathlib import Path
import pandas as pd

# Requer: pip install pandas openpyxl pywin32
# Requer também: Adobe Photoshop instalado no Windows
#
# O script lê uma planilha Excel e atualiza as camadas de texto
# no arquivo PSD do flyer.
#
# Exemplo de planilha:
# camada            texto
# desc-produto-01   CUPIM A MARINADO KG
# desc-produto-02   PERNIL SERRADO KG
# desc-produto-03   MAÇÃ DO PEITO KG
# desc-produto-04   BANANINHA BOVINA KG
# desc-produto-05   CUPIM B KG
# desc-produto-06   PANCETA SUINA KG
# desc-produto-07   FILÉ DE PEITO DE FRANGO KG
# desc-produto-08   LINGUIÇA TOSCANA SEARA KG
# desc-produto-09   PALETA BOVINA KG
#
# Você também pode usar colunas por posição, por exemplo:
# produto_01, produto_02, ..., produto_09
# Nesse caso o script também entende.

try:
    import win32com.client
except ImportError:
    raise SystemExit(
        "pywin32 não está instalado. Rode: pip install pywin32"
    )


PSD_PATH = r"C:\Users\RAFAEL\Documents\PROGRAMAÇÃO\AUTOMATIZA PHOTOSHOP\PSD\9 ITENS AÇOUGUE.psd"
PLANILHA_PATH = r"C:\Users\RAFAEL\Documents\PROGRAMAÇÃO\AUTOMATIZA PHOTOSHOP\CSV\prudutos.xlsx"
ABA_PLANILHA = 0  # pode ser nome da aba, ex: 'Sheet1'
PASTA_SAIDA = r"C:\Users\RAFAEL\Documents\PROGRAMAÇÃO\AUTOMATIZA PHOTOSHOP\RESULTADO"
NOME_BASE_EXPORTACAO = "flyer_atualizado"
EXPORTAR_JPG = True
QUALIDADE_JPG = 10  # 0 a 12
SALVAR_PSD_COPIA = True

# Grupo principal onde estão as camadas, se quiser limitar a busca.
# Se deixar None, ele procura no documento inteiro.
NOME_GRUPO = "09 ITENS copiar 3"

# Mapeamento padrão caso sua planilha use colunas produto_01...produto_09
MAPEAMENTO_COLUNAS = {
    "produto_01": "desc-produto-01",
    "produto_02": "desc-produto-02",
    "produto_03": "desc-produto-03",
    "produto_04": "desc-produto-04",
    "produto_05": "desc-produto-05",
    "produto_06": "desc-produto-06",
    "produto_07": "desc-produto-07",
    "produto_08": "desc-produto-08",
    "produto_09": "desc-produto-09",
}


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def carregar_mapa_textos(arquivo_excel, aba=0):
    df = pd.read_excel(arquivo_excel, sheet_name=aba)
    df.columns = [str(c).strip() for c in df.columns]

    # Formato 1: colunas 'camada' e 'texto'
    if {"camada", "texto"}.issubset(set(df.columns)):
        mapa = {}
        for _, row in df.iterrows():
            camada = normalizar_texto(row["camada"])
            texto = normalizar_texto(row["texto"])
            if camada:
                mapa[camada] = texto
        return mapa

    # Formato 2: primeira linha com colunas produto_01...produto_09
    colunas_encontradas = [c for c in df.columns if c in MAPEAMENTO_COLUNAS]
    if colunas_encontradas:
        primeira_linha = df.iloc[0].to_dict()
        mapa = {}
        for coluna, nome_camada in MAPEAMENTO_COLUNAS.items():
            if coluna in primeira_linha:
                mapa[nome_camada] = normalizar_texto(primeira_linha[coluna])
        return mapa

    raise ValueError(
        "Planilha fora do formato esperado. Use colunas 'camada' e 'texto' "
        "ou colunas 'produto_01' até 'produto_09'."
    )


def conectar_photoshop():
    app = win32com.client.Dispatch("Photoshop.Application")
    app.DisplayDialogs = 3  # psDisplayNoDialogs
    return app


def iterar_camadas(container):
    # ArtLayers do nível atual
    try:
        for i in range(1, container.ArtLayers.Count + 1):
            yield container.ArtLayers.Item(i)
    except Exception:
        pass

    # LayerSets (grupos)
    try:
        for i in range(1, container.LayerSets.Count + 1):
            grupo = container.LayerSets.Item(i)
            yield grupo
            for item in iterar_camadas(grupo):
                yield item
    except Exception:
        pass


def buscar_grupo_por_nome(doc, nome_grupo):
    if not nome_grupo:
        return doc

    for item in iterar_camadas(doc):
        try:
            if hasattr(item, "Name") and item.Name == nome_grupo:
                return item
        except Exception:
            continue

    raise ValueError(f"Grupo '{nome_grupo}' não encontrado no PSD.")


def buscar_camada_texto_por_nome(container, nome_camada):
    for item in iterar_camadas(container):
        try:
            if hasattr(item, "Name") and item.Name == nome_camada:
                # Camada de texto normalmente possui propriedade TextItem
                _ = item.TextItem
                return item
        except Exception:
            continue
    return None


def atualizar_textos(doc, mapa_textos, nome_grupo=None):
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
            print(f"ERRO ao atualizar {nome_camada}: {e}")

    return nao_encontradas


def salvar_psd_como(doc, caminho_saida):
    psd_options = win32com.client.Dispatch("Photoshop.PhotoshopSaveOptions")
    doc.SaveAs(caminho_saida, psd_options, True)


def exportar_jpg(doc, caminho_saida_jpg, qualidade=10):
    jpg_options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
    jpg_options.Quality = qualidade
    jpg_options.EmbedColorProfile = True
    jpg_options.Matte = 1  # sem matte visível
    doc.SaveAs(caminho_saida_jpg, jpg_options, True)


def main():
    psd_path = Path(PSD_PATH)
    planilha_path = Path(PLANILHA_PATH)
    pasta_saida = Path(PASTA_SAIDA)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    if not psd_path.exists():
        raise FileNotFoundError(f"PSD não encontrado: {psd_path}")

    if not planilha_path.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {planilha_path}")

    mapa_textos = carregar_mapa_textos(planilha_path, ABA_PLANILHA)
    print("Mapeamento carregado da planilha:")
    for camada, texto in mapa_textos.items():
        print(f" - {camada}: {texto}")

    app = conectar_photoshop()
    doc = app.Open(str(psd_path.resolve()))

    try:
        nao_encontradas = atualizar_textos(doc, mapa_textos, NOME_GRUPO)

        base = NOME_BASE_EXPORTACAO
        psd_saida = pasta_saida / f"{base}.psd"
        jpg_saida = pasta_saida / f"{base}.jpg"

        if SALVAR_PSD_COPIA:
            salvar_psd_como(doc, str(psd_saida.resolve()))
            print(f"PSD salvo em: {psd_saida}")

        if EXPORTAR_JPG:
            exportar_jpg(doc, str(jpg_saida.resolve()), QUALIDADE_JPG)
            print(f"JPG salvo em: {jpg_saida}")

        if nao_encontradas:
            print("\nCamadas não encontradas:")
            for nome in nao_encontradas:
                print(f" - {nome}")
        else:
            print("\nTodas as camadas foram atualizadas com sucesso.")

    finally:
        # Fecha sem salvar por cima do original.
        # 2 = psDoNotSaveChanges
        doc.Close(2)


if __name__ == "__main__":
    main()
