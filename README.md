# 🛒 Automação de Flyer no Photoshop

Este projeto tem como objetivo automatizar a criação de flyers
promocionais utilizando Python + Adobe Photoshop, eliminando tarefas
manuais repetitivas e acelerando a produção de materiais de divulgação.

## 🚀 O que o projeto faz

Atualmente, o sistema:

-   📄 Lê uma planilha Excel (`.xlsx`) com os produtos
-   ✏️ Atualiza automaticamente os textos no arquivo `.PSD`
-   🏷️ Preenche:
    -   Nome dos produtos
    -   Preços (separados corretamente)
-   🧠 Remove duplicação de preço na descrição automaticamente
-   🖼️ Exporta o flyer final em:
    -   `.PSD` (editável)
    -   `.JPG` (pronto para uso)

------------------------------------------------------------------------

## 🧩 Estrutura do projeto

    automacao-photoshop/
    ├─ CSV/
    │  └─ dados_flyer.xlsx
    ├─ PSD/
    │  └─ 9 ITENS AÇOUGUE copiar.psd
    ├─ automacao_flyer_photoshop.py
    ├─ requirements.txt
    └─ saida/

------------------------------------------------------------------------

## ⚙️ Tecnologias utilizadas

-   Python
-   Pandas
-   PyWin32 (integração com Photoshop)
-   Excel (.xlsx)
-   Adobe Photoshop (COM API)

------------------------------------------------------------------------

## ▶️ Como executar

1.  Instale as dependências:

```{=html}
<!-- -->
```
    pip install pandas openpyxl pywin32

2.  Organize os arquivos:

-   Coloque o `.psd` na pasta `PSD`
-   Coloque o `.xlsx` na pasta `CSV`

3.  Execute:

```{=html}
<!-- -->
```
    python automacao_flyer_photoshop.py

4.  O resultado será gerado na pasta:

```{=html}
<!-- -->
```
    saida/

------------------------------------------------------------------------

## 📊 Formato da planilha

A planilha deve conter as colunas:

  descrição             preco
  --------------------- -------
  BANANINHA BOVINA KG   39,99

------------------------------------------------------------------------

## 💡 Problema resolvido

Antes: - Alteração manual no Photoshop - Alto risco de erro - Processo
lento

Depois: - Automação completa via script - Padronização - Ganho de tempo
significativo ⚡

------------------------------------------------------------------------

## 🔮 Melhorias futuras

-   📅 Automatizar data do flyer
-   💰 Automatizar preços e unidades de medida (KG, UN, etc.)
-   🖼️ Inserir automaticamente imagens dos produtos a partir de um banco
    de dados local
-   🔄 Integração com banco de dados local
-   🌐 Possível integração com API / sistema web

------------------------------------------------------------------------

## 📌 Status do projeto

🚧 Em desenvolvimento contínuo

------------------------------------------------------------------------

## 👨‍💻 Autor

Desenvolvido por **Rafael Ferreira Damasceno**

------------------------------------------------------------------------

## 📄 Licença

Este projeto é livre para estudo e aprimoramento.
