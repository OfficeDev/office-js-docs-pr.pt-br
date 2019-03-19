---
title: Conjuntos de requisitos de API JavaScript do Excel
description: ''
ms.date: 02/15/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 90cf68faaaa7e49d1aa8e77c644ac0ca134420b9
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600309"
---
# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Excel

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Excel são executados em várias versões do Office, incluindo Office 2016 ou posterior para Windows, Office para iPad, Office para Mac e Office Online. A tabela a seguir lista conjuntos de requisitos do Excel, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e as versões ou números de build desses aplicativos.

> [!NOTE]
> Para usar APIs em qualquer um dos conjuntos de requisitos numerados, faça referência à biblioteca **production** no CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Para obter informações sobre o uso de APIs de visualização, confira a seção [APIs de visualização do JavaScript para Excel](#excel-javascript-preview-apis) neste artigo.

|  Conjunto de requisitos  |  Office 365 para Windows  |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Visualização  | Use a versão mais recente do Office para testar as APIs de visualização (talvez seja exigido ser membro do [programa Office Insider](https://products.office.com/office-insider)) |
| ExcelApi1.8  | Versão 1808 (Build 10730.20102) ou posterior | 2.17 ou posterior | 16.17 ou posterior | Setembro de 2018 | Em breve |
| ExcelApi1.7  | Versão 1801 (Build 9001.2171) ou posterior   | 2.9 ou posterior | 16.9 ou posterior | Abril de 2018 | Em breve |
| ExcelApi1.6  | Versão 1704 (Compilação 8201.2001) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Abril de 2017 | Em breve|
| ExcelApi1.5  | Versão 1703 (Compilação 8067.2070) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Março de 2017 | Em breve|
| ExcelApi1.4  | Versão 1701 (build 7870.2024) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Janeiro de 2017 | Em breve|
| ExcelApi1.3  | Versão 1608 (build 7369.2055) ou posterior | 1.27 ou posterior |  15.27 ou posterior| Setembro de 2016 | Versão 1608 (build 7601.6800) ou posterior|
| ExcelApi1.2  | Versão 1601 (build 6741.2088) ou posterior | 1.21 ou posterior | 15.22 ou posterior| janeiro de 2016 ||
| ExcelApi1.1  | Versão 1509 (build 4266.1001) ou posterior | 1.19 ou posterior | 15.20 ou posterior| janeiro de 2016 ||

> [!NOTE]
> O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém o conjunto de requisitos 1.1 de ExcelApi.

Para saber mais sobre as versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários. A tabela a seguir lista as APIs atualmente disponíveis na visualização. Para fornecer feedback sobre uma API de visualização, use o mecanismo de feedback no final da página da Web em que a API está documentada.

> [!NOTE]
> As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Recomendamos que você experimente apenas em ambiente de teste e desenvolvimento. Não use APIs de visualização em um ambiente de produção ou em documentos essenciais aos negócios.
>
> Para usar as APIs de visualização, você deve fazer referência à biblioteca **beta** no CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js e também pode ser necessário ingressar no programa Office Insider para obter uma compilação do Office suficientemente recente.

Atualmente, mais de 400 novas APIs do Excel estão em visualização. A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada. Experimente os novos recursos e dê sua opinião.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Segmentação de Dados | Insira e configure as segmentações de dados em tabelas e Tabelas dinâmicas. | [Segmentação de dados](/javascript/api/excel/excel.slicer) |
| Comentários | Adicione, edite e exclua comentários. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Formas | Inserir, posicionar e formatar imagens, formas geométricas e caixas de texto. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| Novos Gráficos | Explore os novos tipos de gráficos compatíveis: mapas, caixa estreita, cascata, explosão solar, pareto. e funil. | [Chart](/javascript/api/excel/excel.charttype) |
| Filtro automático | Adicionar filtros aos intervalos. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| Áreas | Suporte para intervalos descontínuos. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| Células especiais | Obtenha células que contêm datas, comentários ou fórmulas dentro de um intervalo. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| Encontrar | Encontre valores ou fórmulas em uma planilha ou intervalo. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| Copiar Colar | Copie fórmulas, formatos e valores de um intervalo para outro. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| RangeFormat | Novos recursos com formatos de intervalo. | [Range](/javascript/api/excel/excel.rangeformat) |
| Salvar e fechar pasta de trabalho | Salve e feche a pasta de trabalho.  | [Workbook](/javascript/api/excel/excel.workbook) |
| Inserir pasta de trabalho | Insira uma pasta de trabalho em outra.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Cálculo | Maior controle sobre o mecanismo de cálculo do Excel. | [Aplicativo](/javascript/api/excel/excel.application) |

Veja a seguir uma lista completa das APIs na visualização.

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Retorna um número sobre a versão do Mecanismo de Cálculo do Excel que serviu de base para o recálculo da pasta de trabalho. Somente leitura.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Retorna um CalculationState que indica o estado de cálculo do aplicativo. Para saber detalhes, confira Excel.CalculationState. Somente leitura.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Retorna as configurações do Cálculo iterativo.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Suspende a atualização da tela até que o próximo "context.sync()" seja chamado.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Aplica o Filtro automático em um intervalo e filtra a coluna se o índice de coluna e os critérios de filtro forem especificados.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Limpa os critérios se o Filtro automático tiver filtros|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Retorna um objeto Range que representa o intervalo no qual o Filtro automático se aplica.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Se houver um objeto Range associado ao Filtro automático, esse método o retornará.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|A matriz tem todos os critérios de filtro em um intervalo filtrado automaticamente. Somente Leitura.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Indica se o Filtro automático está ativado ou não. Somente Leitura.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Indica se o Filtro automático tem critérios de filtro. Somente Leitura.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Aplica o objeto Autofilter especificado que está atualmente no intervalo.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Remove o Filtro automático do intervalo.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)||
||[style](/javascript/api/excel/excel.cellborder#style)||
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)||
||[weight](/javascript/api/excel/excel.cellborder#weight)||
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)||
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)||
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)||
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)||
||[left](/javascript/api/excel/excel.cellbordercollection#left)||
||[direita](/javascript/api/excel/excel.cellbordercollection#right)||
||[top](/javascript/api/excel/excel.cellbordercollection#top)||
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)||
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)||
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)||
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)||
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)||
||[padrão](/javascript/api/excel/excel.cellpropertiesfill#pattern)||
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)||
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)||
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)||
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)||
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)||
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)||
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)||
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)||
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)||
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)||
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)||
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)||
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)||
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)||
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)||
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)||
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)||
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)||
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)||
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)||
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)||
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)||
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)||
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)||
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)||
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)||
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Crie e abra uma nova pasta de trabalho.  Opcionalmente, a pasta de trabalho pode ser preenchida com um arquivo. xlsx na base 64.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)||
||[bloqueado](/javascript/api/excel/excel.cellpropertiesprotection#locked)||
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Representa o valor após a alteração. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Representa o valor antes da alteração. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Representa o tipo de valor após a alteração.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Representa o tipo de valor antes da alteração.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Ative o gráfico na interface do usuário do Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Encapsula as opções de gráfico dinâmico. Somente leitura.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Retorna ou define um valor inteiro que representa a esquema de cores do gráfico. Leitura/gravação.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|True se a área do gráfico tiver cantos arredondados. Leitura/gravação.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Retorna ou define se o excedente da lixeira está ativado em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Retorna ou define se o estouro negativo da lixeira está ativado em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[Count](/javascript/api/excel/excel.chartbinoptions#count)|Retorna ou define a contagem da lixeira de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Retorna ou define o valor excedente da lixeira de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[tipo](/javascript/api/excel/excel.chartbinoptions#type)|Retorna ou define o tipo de lixeira de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Retorna ou define o valor do estouro negativo da lixeira de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Retorna ou define o valor da largura da lixeira de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Retorna ou define o tipo de cálculo quartil de um gráfico de caixa estreita. Leitura/gravação.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Retorna ou define se os pontos internos são exibidos em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Retorna ou define se a linha média foi mostrada em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Retorna ou define se o marcador médio foi exibido em um gráfico de caixa estreita. Leitura/gravação.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Retorna ou define se os pontos de exceção são exibidos em um gráfico de caixa estreita. Leitura/gravação.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Representa se deve haver um limite de estilo final para as barras de erros.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Representa as partes da barra de erro a serem incluídas. Para saber detalhes, confira Excel.ChartErrorBarsInclude.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Representa a formatação de ErrorBars do gráfico.|
||[tipo](/javascript/api/excel/excel.charterrorbars#type)|Representa o intervalo marcado como barras de erro. Para saber detalhes, confira Excel.ChartErrorBars.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Representa se deve mostrar barras de erro.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Representa a formatação de linha do gráfico.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Retorna ou define a estratégia de rótulos de mapa de série de um gráfico de mapa de região. Leitura/gravação.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Retorna ou define a área do mapa de série de um gráfico de mapa de região. Leitura/gravação.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Retorna ou define tipo de projeção de série de um gráfico de mapa de região. Leitura/gravação.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Representa se deseja exibir os botões do campo de eixo em um Gráfico dinâmico.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Representa se deseja exibir todos os botões do campo de legenda em um Gráfico dinâmico.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Representa se deseja exibir todos os botões do campo de filtro em um Gráfico dinâmico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Representa se deseja exibir os botões do campo de valor em um Gráfico dinâmico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Retorna ou define o fator de dimensionamento para balões no grupo de gráficos especificado. Pode ser um valor inteiro de 0 (zero) a 300, correspondente a uma porcentagem do tamanho padrão. Aplica-se apenas a gráficos de bolhas. Leitura/gravação.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Retorna ou define a cor para o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Retorna ou define o tipo para o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Retorna ou define o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Retorna ou define a cor para o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Retorna ou define o tipo para o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Retorna ou define o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Retorna ou define a cor para o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Retorna ou define o tipo para o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Retorna ou define o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Retorna ou define o estilo de gradiente da série de um gráfico de mapa da região. Leitura/gravação.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Retorna ou define a cor de preenchimento para pontos de dados negativo de uma série. Leitura/gravação.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Retorna ou define a área de estratégia de rótulo pai de série de um gráficos de mapa de árvore. Leitura/gravação.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Encapsula as opções Lixeira apenas para o gráfico de histograma e gráfico de pareto. Somente leitura.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Encapsula as opções para o gráfico de caixa estreita. Somente leitura.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Encapsula as opções de gráfico de mapa. Somente leitura.|
||[xerrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Representa o objeto de barras de erros para a série de gráficos.|
||[yerrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Representa o objeto de barras de erros para a série de gráficos.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Retorna ou define se as linhas do conector são exibidas em um gráfico de cascata. Leitura/gravação.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|True se o Microsoft Excel mostrar linhas líderes para cada rótulo de dados na série. Leitura/gravação.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Retorna ou define o valor a partir do qual as duas seções de uma pizza ou de uma barra de um gráfico de pizza são separadas. Leitura/gravação.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)||
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)||
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)||
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Obtenha ou defina o conteúdo.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Exclui o thread de comentários. |
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Obtem o local do comentário.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Recebe emails do autor do comentário.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Obtem o nome do autor do comentário.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Obtem a hora de criação do comentário. Retornará nulo se o comentário é convertido como anotação já que, nesse caso, o comentário não terá data de criação.|
||[id](/javascript/api/excel/excel.comment#id)|Representa o identificador de comentário. Somente leitura.|
||[isParent](/javascript/api/excel/excel.comment#isparent)|Representa se é um thread de comentário ou uma resposta. Sempre retornará true aqui. Somente leitura.|
||[replies](/javascript/api/excel/excel.comment#replies)|Representa uma coleção de objetos de resposta associados ao comentário. Somente leitura.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Cria um novo comentário (thread de comentários) com base no conteúdo e local da célula. Um argumento inválido será lançado se o local for maior que uma célula.|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Cria um novo comentário (thread de comentários) com base no conteúdo e local da célula. Um argumento inválido será lançado se o local for maior que uma célula.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtém o número de comentários na coleção.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Retorna um comentário identificado pela respectiva ID. Somente leitura.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtém um comentário com base em sua posição na coleção.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtém um comentário na célula específica na coleção.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtém um comentário relacionado à respectiva ID de resposta na coleção.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.commentcollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Obtenha ou defina o conteúdo.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Exclui a resposta do comentário. |
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obter o local de resposta de comentário.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtém o comentário pai dessa resposta.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Recebe emails do autor da resposta do comentário.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtém o nome do autor da resposta do comentário.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtém a hora de criação de resposta de comentário.|
||[id](/javascript/api/excel/excel.commentreply#id)|Representa o identificador de resposta do comentário. Somente leitura.|
||[isParent](/javascript/api/excel/excel.commentreply#isparent)|Representa se é um thread de comentário ou uma resposta. Sempre retornará false aqui. Somente leitura.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Cria uma resposta de comentário para o comentário.|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Cria uma resposta de comentário para o comentário.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtém o número de respostas de comentários na coleção.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Retorna uma resposta de comentário identificada pela respectiva ID. Somente leitura.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtém uma resposta de comentário com base em sua posição na coleção.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.commentreplycollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Retorna o RangeAreas, compreendendo um ou mais intervalos retangulares, ao qual o formato condicional é aplicado. Somente leitura.|
|[CustomFunctionEventArgs](/javascript/api/excel/excel.customfunctioneventargs)|[higherTicks](/javascript/api/excel/excel.customfunctioneventargs#higherticks)||
||[lowerTicks](/javascript/api/excel/excel.customfunctioneventargs#lowerticks)||
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Retorna um RangeAreas, que consiste em um ou mais intervalos retangulares, com valores inválidos de célula. Se todos os valores de célula forem válidos, essa função gerará um erro ItemNotFound.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Retorna um RangeAreas, que consiste em um ou mais intervalos retangulares, com valores inválidos de célula. Se todos os valores de célula forem válidos, essa função retornará null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|A propriedade usada pelo filtro para realizar a filtragem avançada em richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Retorna o identificador de forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Retorna o objeto de Forma para a forma geométrica. Somente leitura.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Retorna o número de formas no grupo de forma. Somente leitura.|
||[getItem(name: string)](/javascript/api/excel/excel.groupshapecollection#getitem-name-)|Obtém uma forma usando seu respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Obtém uma forma com base em sua posição na coleção.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.groupshapecollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Obtém ou define o rodapé central da planilha.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Obtém ou define o cabeçalho central da planilha.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Obtém ou define o rodapé esquerdo da planilha.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Obtém ou define o cabeçalho esquerdo da planilha.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Obtém ou define o rodapé direito da planilha.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Obtém ou define o cabeçalho direito da planilha.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|O cabeçalho/rodapé geral, usado em todas as páginas, a menos que seja especificada a página par/ímpar ou a primeira página.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|O cabeçalho/rodapé a ser usado para páginas pares, o cabeçalho/rodapé ímpar deve ser especificado para páginas ímpares.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|O cabeçalho/rodapé da primeira página. Para todas as outras páginas, geral ou par/ímpar é usado.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|O cabeçalho/rodapé a ser usado para páginas ímpares, o cabeçalho/rodapé par deve ser especificado para páginas pares.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Obtém ou define o estado do qual os cabeçalhos/rodapés são definidos. Para saber detalhes, confira Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Obtém ou define um sinalizador indicando se os cabeçalhos/rodapés estão alinhados com as margens da página que foram definidas nas opções de layout de página da planilha.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Obtém ou define um sinalizador que indica se os cabeçalhos/rodapés devem ser dimensionados pela escala de porcentagem da página definida nas opções de layout de página da planilha.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Retorna o formato da imagem. Somente leitura.|
||[id](/javascript/api/excel/excel.image#id)|Representa o identificador de forma para o objeto de imagem. Somente leitura.|
||[shape](/javascript/api/excel/excel.image#shape)|Retorna o objeto de forma associado à imagem. Somente leitura.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True se o Excel usará a interação para resolver referências circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Retorna ou define a quantidade máxima de alteração entre cada iteração conforme o Excel resolve referências circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Retorna ou define o número máximo de iterações que o Excel pode usar para resolver uma referência circular.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|Representa o comprimento da ponta da seta no início da linha especificada.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|Representa o estilo da ponta de seta no início da linha especificada.|
||[BeginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|Representa a largura da ponta da seta no início da linha especificada.|
||[connectBeginShape (forma: Excel.Shape, connectionSite: número)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|Conecta o início do conector especificado a uma forma específica.|
||[connectEndShape (forma: Excel.Shape, connectionSite: número)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|Anexa o final do conector especificado a uma forma específica.|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|Representa o tipo de conector de linha.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|Desconecta o início do conector especificado de uma forma.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|Desconecta o final do conector especificado de uma forma.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|Representa o comprimento da ponta de seta no final da linha especificada.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|Representa o estilo da ponta de seta no final da linha especificada.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|Representa a largura da ponta de seta no final da linha especificada.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Representa a forma na qual o início da linha especificada está conectado. Somente leitura.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Representa o site de conexão ao qual o início de um conector está conectado. Somente leitura. Retorna nulo quando o início da linha não está conectado a qualquer forma.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Representa a forma na qual o final da linha especificada está conectado. Somente leitura.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Representa o site de conexão ao qual o final de um conector está conectado. Somente leitura. Retorna nulo quando o final da linha não está conectado a qualquer forma.|
||[id](/javascript/api/excel/excel.line#id)|Representa o identificador de forma. Somente leitura.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Especifica se o início do conector especificado está conectado ou não a uma forma. Somente leitura.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Especifica se o final do conector especificado está conectado ou não a uma forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.line#shape)|Retorna o objeto de forma associado à linha. Somente leitura.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[source](/javascript/api/excel/excel.listdatavalidation#source)|Fonte da lista de validação de dados|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Exclui um objeto de quebra de página.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Obtém a primeira célula após a quebra de página.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Representa o índice de coluna para a quebra de página|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Representa o índice de linha para a quebra de página|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Adiciona uma quebra de página antes da célula superior esquerda do intervalo especificado.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Obtém o número de quebras de página na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Obtém um objeto de quebra de página através do índice.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.pagebreakcollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Redefine todas as quebras de página manuais na coleção.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|Obtém ou define a opção de impressão em preto e branco da planilha.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|Obtém ou define a margem de página inferior da planilha para impressão em pontos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|Obtém ou define o sinalizador de centralização horizontal da planilha. Esse sinalizador determina se a planilha será centralizada horizontalmente quando for impressa.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|Obtém ou define o sinalizador de centralização vertical da planilha. Esse sinalizador determina se a planilha será centralizada verticalmente quando for impressa.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|Obtém ou define a opção de modo de rascunho da planilha. Se for true, a planilha será impressa sem gráficos.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|Obtém ou define o primeiro número de página da planilha a ser impressa. O valor null representa a numeração "automática" de páginas.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|Obtém ou define a margem do rodapé da planilha, em pontos, para usar durante a impressão.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha. Se não houver uma área de impressão, um erro ItemNotFound será gerado.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha. Se não houver uma área de impressão, um objeto null será retornado.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Obtém o objeto range que representa as colunas de título.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Obtém o objeto range que representa as colunas de título. Se não estiver configurado, retornará um objeto null.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Obtém o objeto range representando as linhas do título.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Obtém o objeto range representando as linhas do título. Se não estiver configurado, retornará um objeto null.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|Obtém ou define a margem do cabeçalho da planilha, em pontos, para usar durante a impressão.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|Obtém ou define a margem esquerda da planilha, em pontos, para usar durante a impressão.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|Obtém ou define a orientação de página da planilha.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|Obtém ou define o tamanho do papel da página da planilha.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Obtém ou define se os comentários da planilha deverão ser exibidos durante a impressão.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|Obtém ou define a opção de erros de impressão da planilha.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Obtém ou define um sinalizador de linhas de grade de impressão da planilha. Esse sinalizador determina se as linhas de grade serão impressas ou não.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Obtém ou define um sinalizador de cabeçalhos de impressão da planilha. Esse sinalizador determina se os cabeçalhos serão impressos ou não.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|Obtém ou define a opção de ordem de impressão da página da planilha. Isso especifica a ordem que será usada para processar o número de página impresso.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Configuração de cabeçalho e rodapé da planilha.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|Obtém ou define a margem direita da planilha, em pontos, para usar durante a impressão.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Define a área de impressão da planilha.|
||[setPrintMargins(unit: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Define as margens das páginas da planilha com unidades.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Define as margens das páginas da planilha com unidades.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Define as colunas que contêm as células que serão repetidas à esquerda de cada página da planilha para impressão.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Define as linhas que contêm as células que serão repetidas na parte de cada página da planilha para impressão.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Obtém ou define a margem superior da planilha, em pontos, para usar durante a impressão.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Obtém ou define as opções de zoom de impressão da planilha.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Representa a margem inferior do layout de página na unidade especificada para usar na impressão.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Representa a margem do rodapé do layout de página na unidade especificada para usar na impressão.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Representa a margem do cabeçalho do layout de página na unidade especificada para usar na impressão.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Representa a margem esquerda do layout de página na unidade especificada para usar na impressão.|
||[direita](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Representa a margem direita do layout de página na unidade especificada para usar na impressão.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Representa a margem superior do layout de página na unidade especificada para usar na impressão.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Número de páginas a ser horizontalmente ajustado. Esse valor pode ser null se o dimensionamento por porcentagem for usado.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|O valor do dimensionamento da página de impressão pode estar entre 10 e 400. Esse valor poderá ser null se o ajuste da altura ou largura da página for especificado.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Número de páginas a ser verticalmente ajustado. Esse valor pode ser null se o dimensionamento por porcentagem for usado.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortby: "Ascending" \| "Descending", valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Classifica o Campo dinâmico por valores especificados em um determinado escopo. O escopo define quais valores específicos serão usados na classificação quando|
||[sortByValues(sortby: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Classifica o Campo dinâmico por valores especificados em um determinado escopo. O escopo define quais valores específicos serão usados na classificação quando|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|True se a formatação for formatada automaticamente quando for atualizada ou quando os campos forem movidos|
||[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|True se a lista de campos deve ser mostrada ou ocultada na interface do usuário.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtém a célula no corpo de dados da Tabela dinâmica que contém o valor para a interseção dos objetos dataHierarchy, rowItems e columnItems especificados.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Obtém o DataHierarchy que é usado para calcular o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[getPivotItems(axis: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtém os Itens dinâmicos de um eixo que compõem o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtém os Itens dinâmicos de um eixo que compõem o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|True se a formatação for preservada quando o relatório for atualizado ou recalculado através de operações como dinamização, classificação ou alteração dos itens do campo da página.|
||[setAutosortOnCell(cell: Range \| string, sortby: "Ascending" \| "Descending")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Define uma classificação automática usando a célula especificada para selecionar automaticamente todos os critérios e contextos para a classificação.|
||[setAutosortOnCell(cell: Range \| string, sortby: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Define uma classificação automática usando a célula especificada para selecionar automaticamente todos os critérios e contextos para a classificação.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|True se a tabela dinâmica tiver que usar listas personalizadas na classificação.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|True se a tabela dinâmica tiver que usar listas personalizadas na classificação.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Exclui a Tabela Dinâmica.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Cria uma duplicata desta Tabela Dinâmica com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtém o nome da Tabela Dinâmica.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Verdadeiro significa que esse objeto Tabela Dinâmica é somente leitura. Somente leitura.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Cria uma Tabela Dinâmica em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Obtém o número de estilos de PivotTable na coleção.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Obtém a Tabela Dinâmica padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Obtém um PivotTableStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Obtém um PivotTableStyle por nome. Se PivotTableStyle não existir, retornará um objeto null.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.pivottablestylecollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault (newDefaultStyle: PivotTableStyle \| cadeia de caracteres)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Define a Tabela Dinâmica padrão para uso no escopo do objeto pai.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Os preenchimentos variam do intervalo atual até o intervalo de destino.|
||[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Os preenchimentos variam do intervalo atual até o intervalo de destino.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Converte o intervalo de células com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Converte as células de intervalo em um tipo de dados vinculado na planilha.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o intervalo atual.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o intervalo atual.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Faz o preenchimento relâmpago no intervalo atual. O preenchimento relâmpago preenche automaticamente dados quando detecta um padrão. Portanto, o intervalo deve ser de coluna única e ter dados em torno para encontrar o padrão.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Retorna uma matriz 2D encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada coluna.  Para propriedades que não são consistentes nas células de uma determinada coluna, será retornado null.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.  Para propriedades que não são consistentes nas células de uma determinada linha, será retornado null.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas"\|  "SameConditionalFormat" \| "SameDataValidation" \|  "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \|  "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \|"Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \|"LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \|"Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Falha se aplicado a um intervalo com mais de uma célula. Somente leitura.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Somente leitura.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Falha se aplicado a um intervalo com mais de uma célula. Somente leitura.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Somente leitura.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Obtém uma coleção de tabelas com escopo que se sobrepõe ao intervalo.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Representa se todas as células têm uma borda de despejo.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Representa o estado do tipo de dados de cada célula. Somente leitura.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Remove valores duplicados do intervalo especificado pelas colunas.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados no intervalo atual.|
||[setCellProperties (cellPropertiesData: SettableCellProperties [][]\| OfficeExtension.ClientResult < SettableCellProperties [][]>)](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Atualiza o intervalo com base em uma matriz 2D de propriedades da célula, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[] \| OfficeExtension.ClientResult<SettableColumnProperties[]>)](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Atualiza o intervalo com base em uma única matriz dimensional de propriedades da coluna, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Define um intervalo a ser recalculado quando o próximo recálculo ocorrer.|
||[setRowProperties (rowPropertiesData: SettableRowProperties[] \| OfficeExtension.ClientResult < SettableRowProperties []>)](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Atualiza o intervalo com base em uma única matriz dimensional de propriedades da linha, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Calcula todas as células no RangeAreas.|
||[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Limpa valores, formato, preenchimento, borda, etc. em cada uma das áreas que compõe este objeto RangeAreas.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Limpa valores, formato, preenchimento, borda, etc. em cada uma das áreas que compõe este objeto RangeAreas.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Converte todas as células de RangeAreas com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Converte todas as células de RangeAreas em tipos de dados vinculados.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o RangeAreas atual.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o RangeAreas atual.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Retorna um objeto RangeAreas que representa as colunas inteiras dos objetos RangeAreas (por exemplo, se o RangeAreas atual representa as células "B4:E11, H2", ele retorna um RangeAreas que representa as colunas "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Retorna um objeto RangeAreas que representa as linhas inteiras dos objetos RangeAreas (por exemplo, se o RangeAreas atual representa as células "B4:E11", ele retorna um RangeAreas que representa as linhas "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Retorna o objeto RangeAreas que representa a interseção dos intervalos fornecidos ou RangeAreas. Se nenhuma interseção for encontrada, um erro ItemNotFound será gerado.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Retorna o objeto RangeAreas que representa a interseção dos intervalos fornecidos ou RangeAreas. Se nenhuma interseção for encontrada, um objeto null será retornado.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Retorna um objeto RangeAreas que é deslocado pelo deslocamento de linha e coluna específico. A dimensão do RangeAreas retornado corresponderá ao objeto original. Se o RangeAreas resultante for imposto para fora dos limites da grade da planilha, o sistema gerará um erro.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas"\| "SameConditionalFormat" \| "SameDataValidation" \|  "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \|  "ErrorsLogicalNumber" \| "ErrorsLogicalText" \|  "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados. Gera um erro se nenhuma célula especial que corresponda aos critérios for encontrada.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados. Gera um erro se nenhuma célula especial que corresponda aos critérios for encontrada.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \|"Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \|"LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \|"Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados. Retorna um objeto null se nenhuma célula especial que corresponda ao critério for encontrada.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados. Retorna um objeto null se nenhuma célula especial que corresponda ao critério for encontrada.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Retorna uma coleção de tabelas com escopo que se sobrepõe a qualquer intervalo neste objeto RangeAreas.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Retorna o RangeAreas usado que compreende todas as áreas utilizadas de intervalos retangulares individuais no objeto RangeAreas.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Retorna o RangeAreas usado que compreende todas as áreas utilizadas de intervalos retangulares individuais no objeto RangeAreas.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Retorna a referência RageAreas no estilo A1. O valor do endereço conterá o nome da planilha para cada bloco retangular de células (por exemplo, "Sheet1!A1:B4, Sheet1!D1:D4"). Somente leitura.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Retorna a referência RageAreas na localidade do usuário.  Somente leitura.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Retorna o número de intervalos retangulares que compõem este objeto RangeAreas.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Retorna uma coleção de intervalos retangulares que compõem este objeto RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Retorna o número de células no objeto RangeAreas somando as contagens de células de todos os intervalos retangulares individuais. Retornará -1 se a contagem de células exceder 2^31-1 (2.147.483.647). Somente leitura.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Retorna uma coleção de ConditionalFormats que se cruza com qualquer célula nesse objeto RangeAreas. Somente leitura.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Retorna um objeto dataValidation para todos os intervalos no RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Retorna um objeto rangeFormat encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades para todos os intervalos no objeto RangeAreas. Somente leitura.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Indica se todos os intervalos neste objeto RangeAreas representam colunas inteiras (por exemplo, "A:C, Q:Z"). Somente leitura.|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Indica se todos os intervalos neste objeto RangeAreas representam linhas inteiras (por exemplo, "1:3, 5:7"). Somente leitura.|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Retorna a planilha para o RangeAreas atual. Somente leitura.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Define o RangeAreas que será recalculado quando o próximo recálculo ocorrer.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Representa o estilo de todos os intervalos nesse objeto RangeAreas.|
||[track()](/javascript/api/excel/excel.rangeareas#track--)|Acompanha o objeto para ajuste automático com base nas alterações adjacentes no documento. Essa chamada é uma abreviação de context.trackedObjects.add(thisObject). Se você estiver usando esse objeto em chamadas ".sync" e fora da execução sequencial de um lote ".run" e receber um erro "InvalidObjectPath" ao definir uma propriedade ou invocar um método no objeto, era necessário ter adicionado o objeto à coleção de objetos rastreados quando o objeto foi criado pela primeira vez.|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|Libere a memória associada a este objeto, se ele já tiver sido rastreado anteriormente. Essa chamada é uma abreviação de context.trackedObjects.remove(thisObject). Ter muitos objetos rastreados desacelera o aplicativo host, por isso, lembre-se de liberar todos os objetos adicionados após usá-los. Você precisa chamar "context.sync()" antes da liberação da memória entrar em vigor.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece a cor para a Borda do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece a cor para as Bordas do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Retorna o número de intervalos no RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Retorna o objeto range com base em sua posição no RangeCollection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.rangecollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[padrão](/javascript/api/excel/excel.rangefill#pattern)|Obtém ou define o padrão de um intervalo. Para saber detalhes, confira Excel.FillPattern. LinearGradient e RectangularGradient não são compatíveis.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Define o código de cor HTML que representa a cor do padrão Range, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor padrão para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Representa o status da fonte em tachado. Um valor nulo indica que todo o intervalo não tem configuração de tachado uniforme.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Representa o status da fonte em subscrito.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Representa o status da fonte em sobrescrito.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para a Fonte do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Indica se o texto é automaticamente recuado quando o alinhamento de texto é definido como distribuição igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|A ordem de leitura para o intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Número de linhas duplicadas removidas pela operação.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Número de linhas restantes exclusivas presentes no intervalo resultante.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Especifica se a correspondência deve ser completa ou parcial. O padrão é false (parcial).|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas. O padrão é false (não diferencia maiúsculas de minúsculas).|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)||
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)||
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)||
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Especifica se a correspondência deve ser completa ou parcial. O padrão é false (parcial).|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas. O padrão é false (não diferencia maiúsculas de minúsculas).|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Especifica a direção da pesquisa. O padrão é para frente. Confira Excel.SearchDirection.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)||
||[hiperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)||
||[style](/javascript/api/excel/excel.settablecellproperties#style)||
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)||
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)||
|[Configuração](/javascript/api/excel/excel.setting)|[](/javascript/api/excel/excel.setting#replacestringdatewithdate)||
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Retorna ou define o texto da descrição alternativa de um objeto de forma.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Retorna ou define o texto do título alternativo de um objeto de forma.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Remove a forma da planilha.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Representa o tipo de forma geométricas da forma geométrica. Para saber detalhes, confira Excel.GeometricShapeType. Retorna nulo se o tipo de forma não for "GeometricShape".|
||[getAsImage(format: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#getasimage-format-)|Converte a forma em uma imagem e retorna a imagem como uma cadeia de caracteres de base 64. O DPI é 96. Os formatos com suporte apenas são `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, e `Excel.PictureFormat.GIF`.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Converte a forma em uma imagem e retorna a imagem como uma cadeia de caracteres de base 64. O DPI é 96. Os formatos com suporte apenas são `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, e `Excel.PictureFormat.GIF`.|
||[height](/javascript/api/excel/excel.shape#height)|Representa a altura, em pontos, da forma.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Move a forma horizontalmente pelo número especificado de pontos.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|O formato é girado em sentido horário ao redor do eixo z pelo número especificado de graus.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Move a forma verticalmente pelo número especificado de pontos.|
||[left](/javascript/api/excel/excel.shape#left)|A distância, em pontos, da lateral esquerda da forma do lado  esquerdo da planilha.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Especifica se a taxa de proporção dessa forma está bloqueada ou não.|
||[name](/javascript/api/excel/excel.shape#name)|Representa o nome da forma.|
||[placement](/javascript/api/excel/excel.shape#placement)|Representa como o objeto é anexado às células abaixo dela.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Retorna o número de locais de conexão nessa forma. Somente leitura.|
||[fill](/javascript/api/excel/excel.shape#fill)|Retorna a formatação de preenchimento dessa forma. Somente leitura.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Retorna a forma geométrica associada à forma. Um erro será lançado, se o tipo de forma não for "GeometricShape".|
||[group](/javascript/api/excel/excel.shape#group)|Retorna o grupo de forma associado à forma. Um erro será lançado, se o tipo de forma não for "GroupShape".|
||[id](/javascript/api/excel/excel.shape#id)|Representa o identificador de forma. Somente leitura.|
||[image](/javascript/api/excel/excel.shape#image)|Retorna a imagem associada à forma. Um erro será lançado, se o tipo de forma não for "Imagem".|
||[level](/javascript/api/excel/excel.shape#level)|Representa o nível da forma especificada. Por exemplo, um nível de 0 significa que a forma não faz parte de nenhum grupo, um nível de 1 significa que a forma é parte de um grupo de nível superior e um nível 2 significa que a forma faz parte de um subgrupo do nível superior.|
||[line](/javascript/api/excel/excel.shape#line)|Retorna a linha associada à forma. Um erro será lançado, se o tipo de forma não for "Linha".|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Retorna a formatação de linha do objeto de forma. Somente leitura.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Ocorre quando a forma é ativada.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Ocorre quando a forma é desativada.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Representa o grupo pai dessa forma.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Retorna o objeto text frame de uma forma. Somente leitura.|
||[tipo](/javascript/api/excel/excel.shape#type)|Retorna o tipo dessa forma. Para saber detalhes, confira Excel.ShapeType. Somente leitura.|
||[zorderPosition](/javascript/api/excel/excel.shape#zorderposition)|Retorna a posição da forma especificada na ordem z, com 0 representando a parte inferior da pilha do pedido. Somente leitura.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Representa a rotação, em graus, da forma.|
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Dimensiona a altura da forma por um fator especificado. Para imagens, é possível indicar se você deseja dimensionar a forma em relação ao tamanho original ou ao tamanho atual. As formas que não são figuras serão sempre dimensionadas em relação à sua altura atual.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Dimensiona a altura da forma por um fator especificado. Para imagens, é possível indicar se você deseja dimensionar a forma em relação ao tamanho original ou ao tamanho atual. As formas que não são figuras serão sempre dimensionadas em relação à sua altura atual.|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Dimensiona a largura da forma por um fator especificado. Para imagens, é possível indicar se você deseja dimensionar a forma em relação ao tamanho original ou ao tamanho atual. As formas que não são figuras serão sempre dimensionadas em relação à sua largura atual.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Dimensiona a largura da forma por um fator especificado. Para imagens, é possível indicar se você deseja dimensionar a forma em relação ao tamanho original ou ao tamanho atual. As formas que não são figuras serão sempre dimensionadas em relação à sua largura atual.|
||[setZOrder(position: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-position-)|Move a forma especificada para cima ou para baixo na ordem z da coleção, que a desloca para frente ou para trás de outras formas.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Move a forma especificada para cima ou para baixo na ordem z da coleção, que a desloca para frente ou para trás de outras formas.|
||[top](/javascript/api/excel/excel.shape#top)|A distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[visible](/javascript/api/excel/excel.shape#visible)|Representa a visibilidade essa forma.|
||[width](/javascript/api/excel/excel.shape#width)|Representa a largura, em pontos, da forma.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Obtém o id da forma ativada.|
||[tipo](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Obtém a id da planilha na qual a forma está ativada.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus”)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Adiciona uma forma geométrica à planilha. Retorna um objeto Shape que representa a nova forma.|
||[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Adiciona uma forma geométrica à planilha. Retorna um objeto Shape que representa a nova forma.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Um subconjunto de formas na planilha do conjunto de grupos. Retorna um objeto Shape que representa o novo grupo de formas.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Cria uma imagem de uma cadeia de caracteres na base 64 e a adiciona à planilha. Retorna o objeto Shape que representa a nova imagem.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adiciona uma linha à planilha. Retorna um objeto Shape que representa a nova linha.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adiciona uma linha à planilha. Retorna um objeto Shape que representa a nova linha.|
||[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha. Retorna um objeto Shape que representa a nova imagem.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Adiciona uma caixa de texto na planilha com o texto fornecido como conteúdo. Retorna um objeto Shape que representa a nova caixa de texto.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Retorna o número de formas da planilha. Somente leitura.|
||[getItem(name: string)](/javascript/api/excel/excel.shapecollection#getitem-name-)|Obtém uma forma usando seu respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Obtém uma forma usando sua posição na coleção.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.shapecollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Obtém o id da forma que está desativada.|
||[tipo](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Obtém a id da planilha na qual a forma está desativada.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Limpa a formatação do preenchimento de um objeto de forma.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Representa o primeiro plano de preenchimento da forma para cor no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[tipo](/javascript/api/excel/excel.shapefill#type)|Retorna o tipo de preenchimento da forma. Somente leitura. Para saber detalhes, confira Excel.ShapeFillType.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Define a formatação de preenchimento de um formato com uma cor uniforme. Isso altera o tipo de preenchimento para "Sólido".|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Retorna ou define a porcentagem de transparência do preenchimento especificado como um valor de 0,0 (opaco) a 1,0 (transparente). Retorna nulo se o tipo de forma não suportar transparência ou se o preenchimento de forma tiver transparência inconsistente como com um tipo de preenchimento de gradiente.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Representa o status da fonte em negrito. Retornará null se o TextRange incluir fragmentos de texto em negrito e não em negrito.|
||[color](/javascript/api/excel/excel.shapefont#color)|A representação de código de cor HTML para a cor do texto. (Por exemplo, #FF0000 representa vermelho). Retornará null se o TextRange incluir fragmentos de texto com cores diferentes.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Representa o status da fonte em itálico. Retorna null se o TextRange incluir fragmentos de texto em itálico e que não está em itálico.|
||[name](/javascript/api/excel/excel.shapefont#name)|Representa o nome da fonte (por exemplo, "Calibri"). Se o texto estiver no idioma Script Complexo ou Leste Asiático, esse é o nome da fonte correspondente. Caso contrário, esse é o nome da fonte Latin.|
||[size](/javascript/api/excel/excel.shapefont#size)|Representa o tamanho da fonte em pontos (por exemplo, 11). Retorna nulo se o TextRange incluir fragmentos de texto com tamanhos de fontes diferentes.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Tipo de sublinhado aplicado à fonte. Retorna nulo se o TextRange incluir fragmentos de texto com estilos de sublinhado diferentes. Para saber detalhes, confira Excel.ShapeFontUnderlineStyle.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Representa o identificador de forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Retorna o objeto de forma associado ao grupo. Somente leitura.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Retorna uma coleção de objetos de forma. Somente leitura.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Desagrupa todas as formas agrupadas no grupo de forma especificado.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Representa a cor da linha no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos de traços inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Representa o grau de transparência da linha especificada como um valor de 0,0 (opaco) a 1,0 (claro). Retorna nulo quando a forma possui transparências inconsistentes.|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Representa se a formatação de linha de um elemento de forma é visível ou não. Retorna nulo quando a forma possui visibilidades inconsistentes.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Representa a espessura da linha, em pontos. Retorna nulo quando não a linha não estiver visível ou existirem espessuras de linha inconsistentes.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Representa a legenda da segmentação de dados.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Limpa todos os filtros aplicados à segmentação de dados no momento.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Exclui a segmentação de dados.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Retorna uma matriz de chaves de itens selecionados. Somente leitura.|
||[height](/javascript/api/excel/excel.slicer#height)|Representa a altura, em pontos, da segmentação de dados.|
||[left](/javascript/api/excel/excel.slicer#left)|Representa a distância, em pontos, da lateral esquerda da segmentação de dados à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicer#name)|Representa o nome da segmentação de dados.|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Representa o nome usado na fórmula.|
||[id](/javascript/api/excel/excel.slicer#id)|Representa a id exclusiva da segmentação de dados. Somente leitura.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|True se todos os filtros atualmente aplicados à segmentação de dados estiverem desmarcados.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Representa a coleção de SlicerItems que faz parte da segmentação de dados. Somente leitura.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Representa a planilha que contém a segmentação de dados. Somente leitura.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Seleciona os itens da segmentação de dados com base em suas chaves. A seleção anterior será limpa.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Representa a ordem de classificação dos itens na segmentação de dados. Valores possíveis são: DataSourceOrder, Ordem crescente, Ordem decrescente.|
||[style](/javascript/api/excel/excel.slicer#style)|Valor da constante que representa o estilo da Segmentação de dados. Os valores possíveis são: SlicerStyleLight1 thru SlicerStyleLight6, TableStyleOther1 thru TableStyleOther2, SlicerStyleDark1 thru SlicerStyleDark6. Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
||[top](/javascript/api/excel/excel.slicer#top)|Representa a distância, em pontos, da borda superior da segmentação de dados na parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicer#width)|Representa a largura, em pontos, da segmentação de dados.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Adiciona uma nova segmentação de dados à pasta de trabalho.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Retorna o número de segmentações de dados na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Obtém um objeto de segmentação de dados usando seu respectivo nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Obtém uma segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Obtém uma segmentação de dados usando seu nome ou id. Se a ela não existir, retornará um objeto null.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.slicercollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True se o item da segmentação de dados estiver selecionado.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True se o item de segmentação de dados tiver dados.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Representa o valor exclusivo que representa o item da segmentação de dados.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Representa o valor exibido na interface do usuário.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Retorna o número de itens da segmentação de dados na segmentação de dados.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtém um objeto de item da segmentação de dados usando sua chave ou nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtém um item da segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtém um item da segmentação de dados usando sua chave ou nome. Se o item da segmentação de dados não existir, retornará um objeto null.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.sliceritemcollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Exclui o SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Cria uma duplicata deste SlicerStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtém o nome o SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Verdadeiro significa que esse objeto SlicerStyle é somente leitura. Somente leitura.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Cria um SlicerStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Obtém o número de segmentação de estilos na coleção.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Obtém o padrão SlicerStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Obtém uma SlicerStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Obtém uma SlicerStyle por nome. Se o SlicerStyle não existir, retornará um objeto null.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.slicerstylecollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Define o padrão SlicerStyle para uso no escopo do objeto pai.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Representa o subcampo que é o nome da propriedade de destino de um valor avançado para classificação.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Obtém o número de estilos na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Obtém um estilo com base em sua posição na coleção.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Altera a tabela para usar o estilo de tabela padrão.|
||[autoFilter](/javascript/api/excel/excel.table#autofilter)|Representa o objeto AutoFilter da tabela. Somente Leitura.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela específica.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Obtém a id da tabela que é adicionada.|
||[tipo](/javascript/api/excel/excel.tableaddedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Obtém a id da planilha na qual o gráfico é adicionado.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[detalhes](/javascript/api/excel/excel.tablechangedeventargs#details)|Representa informações sobre os detalhes da alteração|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Ocorre quando uma nova tabela é adicionada na pasta de trabalho.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Ocorre quando a tabela especificada é excluída em uma pasta de trabalho.|
||[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela localizada em uma pasta de trabalho ou em uma planilha.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Especifica a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Especifica a id da tabela que é excluída.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Especifica o nome da tabela que é excluída.|
||[tipo](/javascript/api/excel/excel.tabledeletedeventargs#type)|Especifica o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Obtém a id da planilha na qual a tabela é excluída.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Representa a id da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Representa o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Representa a id da planilha que contém a tabela.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Obtém o número de tabelas na coleção.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Obtém a primeira tabela na coleção. As tabelas na coleção são classificadas de cima para baixo e da esquerda para a direita, de forma que a tabela superior esquerda seja a primeira tabela da coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Obtém uma tabela pelo nome ou ID.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.tablescopedcollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Exclui o TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Cria uma duplicata deste TableStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtém o nome do TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Verdadeiro significa que esse objeto TableStyle é somente leitura. Somente leitura.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Cria um TableStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Obtém o número de estilos de tabelas na coleção.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Obtém o padrão TableStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Obtém um TableStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Obtém um TableStyle por nome. Se o TableStyle não existir, retornará um objeto null.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.tablestylecollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Define a TableStyle padrão para uso no escopo do objeto pai..|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|Obtém ou define as configurações de dimensionamento automático para o quadro de texto. Um quadro de texto pode ser configurado para ajustar automaticamente o texto ao quadro de texto, para ajustar automaticamente o quadro do texto ao texto ou não executar qualquer dimensionamento automático.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Exclui todo o texto no quadro de texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Representa o alinhamento horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalAlignment para obter detalhes.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Representa o comportamento de excedente horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalOverflow para obter detalhes.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Representa a orientação do texto do quadro de texto. Confira Excel.ShapeTextOrientation para obter detalhes.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Representa a ordem de leitura do quadro de texto, da direita para a esquerda ou da direita para a esquerda. Confira Excel.ShapeTextReadingOrder para obter detalhes.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Especifica se o quadro de texto contém texto.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Representa o texto que está anexado a uma forma, bem como propriedades e métodos para manipular o texto. Confira Excel.TextRange para obter detalhes.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Representa margem direita, em pontos, do quadro de texto.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Representa o alinhamento vertical do quadro de texto. Confira Excel.ShapeTextVerticalAlignment para obter detalhes.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Representa o comportamento de excedente vertical do quadro de texto. Confira Excel.ShapeTextVerticalOverflow para obter detalhes.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Retorna um objeto TextRange para a subcadeia de caracteres no intervalo especificado.|
||[font](/javascript/api/excel/excel.textrange#font)|Retorna um objeto ShapeFont que representa os atributos de fonte do intervalo de texto. Somente leitura.|
||[text](/javascript/api/excel/excel.textrange#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Exclui o TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Cria uma duplicata deste TimelineStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtém o nome do TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Verdadeiro significa que esse objeto TimelineStyle é somente leitura. Somente leitura.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Cria um TimelineStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Obtém o número de estilos de linha do tempo na coleção.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Obtém o padrão TimelineStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Obtém uma TimelineStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Obtém uma TimelineStyle por nome. Se o TimelineStyle não existir, retornará um objeto null.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.timelinestylecollection#load-option-)|Coloca um comando na fila para carregar as propriedades especificadas do objeto. Você deve chamar "context.sync()" antes de ler as propriedades.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Define o padrão TimelineStyle para uso no escopo do objeto pai.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True se todos os gráficos na pasta de trabalho estiverem rastreando os pontos de dados reais aos quais eles estão anexados.|
||[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fechar a pasta de trabalho atual.|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fechar a pasta de trabalho atual.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Obtém o gráfico ativo no momento na pasta de trabalho. Se não houver um gráfico ativo, será lançada uma exceção quando essa instrução for invocada|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Obtém o gráfico ativo no momento na pasta de trabalho. Se não houver um gráfico ativo, um objeto null será retornado|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtém a segmentação de dados ativa no momento na pasta de trabalho. Se não houver uma segmentação de dados ativa, será lançada uma exceção quando essa instrução for invocada.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtém a segmentação de dados ativa no momento na pasta de trabalho. Se não houver uma segmentação de dados ativa, um objeto null será retornado|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|True se a pasta de trabalho estiver sendo editada por vários usuários (coautoria).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Obtém um ou mais intervalos atualmente selecionados da pasta de trabalho. Ao contrário de getSelectedRange(), esse método retorna um objeto RangeAreas que representa todos os intervalos selecionados.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Especifica se as alterações foram feitas ou não desde que a pasta de trabalho foi salva pela última vez.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Especifica se a pasta de trabalho está ou não no modo de salvamento automático. Somente Leitura.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Retorna um número sobre a versão do Mecanismo de Cálculo do Excel. Somente Leitura.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Representa uma coleção de comentários associados à pasta de trabalho. Somente leitura.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Ocorre quando a configuração Salvamento automático é alterada na pasta de trabalho.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Especifica se a pasta de trabalho já foi salva localmente ou online. Somente Leitura.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Representa uma coleção de SlicerStyles associados à pasta de trabalho. Somente leitura.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Representa uma coleção de segmentações de dados associados à pasta de trabalho. Somente leitura.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Representa uma coleção de TableStyles associadas à pasta de trabalho. Somente leitura.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Representa uma coleção de TimelineStyles associados à pasta de trabalho. Somente leitura.|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|Salvar a pasta de trabalho atual.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Salvar a pasta de trabalho atual.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True se os cálculos dessa pasta de trabalho forem efetuados usando apenas a precisão dos números conforme forem exibidos.|
|[WorkbookAutoSaveSetting [...]](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[tipo](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Representa o tipo do evento. Para saber detalhes, confira Excel.EventType.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Obtém ou define a propriedade enableCalculation da planilha.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Encontra todas as ocorrências de determinada cadeia de caracteres com base nos critérios especificados e as retorna como um objeto RangeAreas, compreendendo um ou mais intervalos retangulares.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Encontra todas as ocorrências de determinada cadeia de caracteres com base nos critérios especificados e as retorna como um objeto RangeAreas, compreendendo um ou mais intervalos retangulares.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Obtém o objeto RangeAreas que representa um ou mais blocos de intervalos retangulares especificados pelo endereço ou nome.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Representa o objeto AutoFilter da planilha. Somente Leitura.|
||[comments](/javascript/api/excel/excel.worksheet#comments)|Retorna um conjunto de todos os objetos Comments na planilha. Somente leitura.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Obtém a coleção de quebra de página horizontal da planilha. Esta coleção contém apenas quebras de página manuais.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Ocorre quando o filtro é aplicado em uma planilha específica.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Ocorre quando o formato é alterado em uma planilha específica.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Obtém o objeto PageLayout da planilha.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Retorna a coleção de todos os objetos Shape na planilha. Somente leitura.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Retorna uma coleção de segmentações de dados que fazem parte da planilha. Somente leitura.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Obtém a coleção de quebra de página vertical da planilha. Esta coleção contém apenas quebras de página manuais.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados na planilha atual.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[detalhes](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Representa informações sobre os detalhes da alteração|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Ocorre quando uma planilha da pasta de trabalho é alterada.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Ocorre quando uma planilha na pasta de trabalho tem o formato alterado.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Ocorre quando a seleção é alterada em uma planilha.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Representa o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Representa a id da planilha na qual o filtro é aplicado.|
|[WorksheetFormatChanged [...]](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica. Pode retornar o objeto null.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Especifica se a correspondência deve ser completa ou parcial. O padrão é false (parcial).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas. O padrão é false (não diferencia maiúsculas de minúsculas).|

## <a name="whats-new-in-excel-javascript-api-18"></a>Quais são as novidades na API JavaScript do Excel 1.8

O conjunto de requisitos 1.8 da API JavaScript do Excel inclui APIs para tabelas dinâmicas, validação de dados, gráficos, eventos de gráficos, opções de desempenho e criação de pasta de trabalho.

### <a name="pivottable"></a>Tabela Dinâmica

Onda 2 das APIs de Tabela Dinâmica permite que os suplementos definam as hierarquias de uma Tabela Dinâmica. Agora você pode controlar os dados e como eles são agregados. Nosso [Artigo de Tabela Dinâmica](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables) tem mais informações sobre a nova funcionalidade de tabela dinâmica.

### <a name="data-validation"></a>Validação de Dados

A validação de dados permite controlar o que um usuário digita em uma planilha. Você pode limitar as células a conjuntos de respostas predefinidos ou fornecer avisos pop-up sobre entradas indesejadas. Saiba mais sobre [adicionar a validação de dados para intervalos](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation) hoje.

### <a name="charts"></a>Gráficos

Outra rodada de APIs de gráficos traz um controle programático ainda maior sobre os elementos do gráfico. Agora você tem maior acesso à legenda, eixos, linha de tendência e área de plotagem.

### <a name="events"></a>Eventos

Mais [eventos](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events) foram adicionados para os gráficos. Faça o seu suplemento reagir aos usuários interagindo com o gráfico. Você também pode [alternar eventos](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events) disparados em toda a pasta de trabalho.

|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Método_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|Cria uma nova pasta de trabalho oculta usando um arquivo .xlsx com codificação base64 opcional.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Propriedade_ > formula1|Obtém ou define a Formula1, por exemplo, o valor mínimo ou valor, dependendo do operador.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Propriedade_ > formula2|Obtém ou define a Formula2, por exemplo, o valor máximo ou valor, dependendo do operador.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Relação_ > operator|O operador a ser usado para validar os dados.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > categoryLabelLevel|Retorna ou define uma constante de enumeração ChartCategoryLabelLevel referindo-se ao nível de onde os rótulos de categoria estão sendo originados. Leitura/gravação.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > plotVisibleOnly|Verdadeiro se apenas as células visíveis forem plotadas. Falso se ambas as células visíveis e ocultas forem plotadas.. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > seriesNameLevel|Retorna ou define uma constante de enumeração ChartSeriesNameLevel referente ao nível de origem dos nomes das séries. Leitura/gravação.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > showDataLabelsOverMaximum|Representa se os rótulos de dados devem ser mostrados quando o valor for maior que o valor máximo no eixo de valor.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > style|Retorna ou define o estilo do gráfico para o gráfico. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > displayBlanksAs|Retorna ou define a maneira como as células em branco são plotadas em um gráfico. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > plotArea|Representa a plotArea para o gráfico. Somente leitura.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > plotBy|Retorna ou define como as colunas ou linhas são usadas como séries de dados no gráfico. ReadWrite.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriedade_ > chartId|Obtém o id do gráfico que está ativado.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual o gráfico é ativado.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriedade_ > chartId|Obtém o id do gráfico que é adicionado à planilha.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual o gráfico é adicionado.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Relação_ > source|Obtém a origem do evento.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > isBetweenCategories|Representa se o eixo de valor cruza o eixo de categoria entre as categorias.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > multiLevel|Representa se um eixo é multinível ou não.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > numberFormat|Representa o código de formato para o rótulo de marcação do eixo.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > offset|Representa a distância entre os níveis de rótulos e a distância entre o primeiro nível e a linha do eixo. O valor deve ser um inteiro de 0 a 1000.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > positionAt|Representa a posição do eixo especificada onde o outro eixo cruza. Você deve usar o método SetPositionAt (double) para definir essa propriedade. Somente leitura.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > textOrientation|Representa a orientação do texto do rótulo de seleção do eixo. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > alignment|Representa o alinhamento para o rótulo de escala do eixo especificado.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > position|Representa a posição do eixo especificada onde o outro eixo cruza.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|Define a posição do eixo especificada onde o outro eixo cruza.|1,8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_Relação_ > fill|Representa a formatação de preenchimento de gráfico. Somente leitura.|1,8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_Método_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|Um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|1,8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relação_ > border|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|1,8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relação_ > fill|Representa a formatação de preenchimento de gráfico. Somente leitura.|1,8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Método_ > [clear()](/javascript/api/excel/excel.chartborder)|Limpa a formatação da borda de um elemento do gráfico.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > AutoText|Valor booliano que representa se o rótulo de dados gerará automaticamente o texto apropriado com base no contexto..|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > formula|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de dados usando a notação no estilo A1.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > height|Retorna a altura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível. Somente leitura.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > left|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > numberFormat|Valor de cadeia de caracteres que representa o código do formato do rótulo de dados.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > text|Cadeia de caracteres que representa o texto do rótulo de dados em um gráfico.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > textOrientation|Representa a orientação de texto de rótulo de dados do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > top|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > width|Retorna a largura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível. Somente leitura.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relação_ > format|Representa o formato do rótulo de dados do gráfico. Somente leitura.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relação_ > horizontalAlignment|Representa o alinhamento horizontal de rótulo de dados do gráfico.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relação_ > verticalAlignment|Representa o alinhamento vertical do rótulo de dados do gráfico.|1,8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_Relação_ > border|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriedade_ > AutoText|Indica se os rótulos de dados geram automaticamente texto apropriado com base no contexto.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriedade_ > numberFormat|Representa o código de formatação para rótulos de dados.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriedade_ > textOrientation|Representa a orientação de texto dos rótulos de dados. O valor deve ser um número inteiro de -90 a 90 ou 0 o 180 para texto orientado verticalmente.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relação_ > horizontalAlignment|Representa o alinhamento horizontal de rótulo de dados do gráfico.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relação_ > verticalAlignment|Representa o alinhamento vertical do rótulo de dados do gráfico.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriedade_ > chartId|Obtém o id do gráfico que está desativado.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha em que o gráfico está desativado.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriedade_ > chartId|Obtém o id do gráfico que é excluído da planilha.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual o gráfico foi deletado.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Relação_ > source|Obtém a origem do evento.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > height|Representa a altura de legendEntry na legenda do gráfico. Somente leitura.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > index|Representa o índice de legendEntry na legenda do gráfico. Somente leitura.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > left|Representa a esquerda de um gráfico legendEntry. Somente leitura.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > top|Representa a parte superior de um gráfico legendEntry. Somente leitura.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > width|Representa a largura de legendEntry na legenda do gráfico. Somente leitura.|1,8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_Relação_ > border|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > height|Representa o valor de altura de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > insideHeight|Representa o valor insideHeight plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > insideLeft|Representa o valor insideLeft de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > insideTop|Representa o valor insideTop de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > insideWidth|Representa o valor insideWidth de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > left|Representa o valor de plotArea à esquerda.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > top|Representa o valor máximo de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > width|Representa o valor de largura de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relação_ > format|Representa a formatação de um gráfico plotArea. Somente leitura.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relação_ > position|Represente a posição de plotArea.|1,8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relação_ > border|Representa os atributos de borda de um gráfico plotArea. Somente leitura.|1,8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relação_ > fill|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo. Somente leitura.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > explosion|Retorna ou define o valor de explosão para um gráfico de pizza ou fatia de gráfico de rosca. Retorna 0 (zero) se não houver explosão (a ponta da fatia está no centro da pizza). ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > firstSliceAngle|Retorna ou define o ângulo do primeiro gráfico de pizza ou fatia de gráfico de rosca, em graus (no sentido horário a partir da vertical). Aplica-se apenas a pizza, torta 3-D e gráficos de rosca.. Pode ser um valor de 0 a 360. ReadWrite|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > invertIfNegative|Verdadeiro se o Microsoft Excel inverte o padrão no item quando ele corresponde a um número negativo. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > overlap|Especifica como barras e colunas são posicionadas. Pode ser um valor entre -100 e 100. Se aplicam apenas às barras 2D e gráficos de colunas 2D. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > secondPlotSize|Retorna ou define o tamanho da seção secundária de uma pizza do gráfico de pizza ou de uma barra de gráfico de pizza, como uma porcentagem do tamanho da pizza primária. Pode ser um valor de 5 de 200. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > varyByCategories|Verdadeiro se o Microsoft Excel atribuir uma cor ou padrão diferente a cada marcador de dados. O gráfico deve conter apenas uma série. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > axisGroup|Retorna ou define o grupo para a série especificada.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > dataLabels|Representa uma coleção de todos os dataLabels da série. Somente leitura.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > splitType|Retorna ou define a maneira como as duas seções de uma pizza do gráfico de pizza ou de uma barra do gráfico de pizza são divididas. ReadWrite.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > backwardPeriod|Representa o número de períodos que a linha de tendência se estende para trás.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > forwardPeriod|Representa o número de períodos que a linha de tendência se estende para frente.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > showEquation|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > showRSquared|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relação_ > label|Representa o rótulo de linha de tendência um gráfico. Somente leitura.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > AutoText|Valor booliano que representa se o rótulo de linha de tendência gerará automaticamente o texto apropriado com base no contexto..|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > formula|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de linha de tendência usando a notação no estilo A1.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > height|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível. Somente leitura.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > left|Representa a distância, em pontos, da borda esquerda do rótulo da linha de tendência do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > numberFormat|Valor de cadeia de caracteres que representa o código do formato do rótulo de linha de tendência.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > text|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > textOrientation|Representa a orientação de texto de rótulo de linha de tendência de gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > top|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > width|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível. Somente leitura.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relação_ > format|Representa o formato do rótulo de linha de tendência de gráfico. Somente leitura.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relação_ > horizontalAlignment|Representa o alinhamento horizontal de rótulo de linha de tendência de gráfico.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relação_ > verticalAlignment|Representa o alinhamento vertical do rótulo de linha de tendência de gráfico.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relação_ > border|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relação_ > fill|Representa o formato de preenchimento do rótulo de linha de tendência atual do gráfico. Somente leitura.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relação_ > font|Representa os atributos de fonte do rótulo de linha de tendência do gráfico, como nome, tamanho, cor, dentre outros. Somente leitura.|1,8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_Propriedade_ > formula| Uma fórmula de validação de dados personalizados. Isso cria regras especiais de entrada, como impedir duplicatas ou limitar o total em um intervalo de células.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriedade_ > id|ID do DataPivotHierarchy. Somente leitura.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriedade_ > nome|Nome da DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriedade_ > numberFormat|Formato de número do DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriedade_ > posição|Posição da DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relação_ > field|Retorna PivotFields associados a DataPivotHierarchy. Somente leitura.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relação_ > showAs|Determina se os dados devem ser mostrados como um cálculo de resumo específico ou não.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relação_ > summarizeBy|Determina se deve mostrar todos os itens a DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Método_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Redefina a DataPivotHierarchy para os valores padrão.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Propriedade_ > itens|Um conjunto de objetos dataPivotHierarchy. Somente leitura.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém DataPivotHierarchy por nome ou id.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém uma DataPivotHierarchy por nome. Se o DataPivotHierarchy não existir, retornará um objeto nulo.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Propriedade_ > ignoreBlanks|Ignora espaços em branco: nenhuma validação de dados será executada em células vazias, o padrão será definido como verdadeiro.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Propriedade_ > valid|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados. Somente leitura.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relação_ > errorAlert|Alerta de erro quando o usuário insere dados inválidos.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relação_ > prompt|Avisa quando os usuários selecionam uma célula.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relação_ > rule|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relação_ > type|Tipo de validação de dados, confira [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) para obter detalhes. Somente leitura.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Método_ > [clear()](/javascript/api/excel/excel.datavalidation)|Desfazer a validação de dados do intervalo atual.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriedade_ > mensagem|Representa a mensagem de alerta de erro.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriedade_ > showAlert|Determina se deseja mostrar uma caixa de diálogo de alerta de erro ou não quando um usuário insere dados inválidos. O padrão é verdadeiro.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriedade_ > title|Representa o título da caixa de diálogo de alerta de erro.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Relação_ > style|Representa o tipo de alerta de validação de dados, confira [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) para obter detalhes.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).|_Propriedade_ > mensagem|Representa a mensagem a solicitação.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).|_Propriedade_ > showPrompt|Determina se deseja ou não mostrar o prompt quando o usuário seleciona uma célula com a validação de dados.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).|_Propriedade_ > title|Representa o título para a solicitação.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > custom|Critérios de validação de dados personalizados.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > date|Critérios de validação de dados de data.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > decimal|Critérios de validação de dados decimais.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > list|Critérios de validação de dados da lista.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > textLength|Critérios de validação de dados TextLength.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > time|Critérios de validação de dados de tempo.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > wholeNumber|Critérios de validação de dados WholeNumber.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Propriedade_ > formula1|Obtém ou define a Formula1, por exemplo, o valor mínimo ou valor, dependendo do operador.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Propriedade_ > formula2|Obtém ou define a Formula2, por exemplo, o valor máximo ou valor, dependendo do operador.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Relação_ > operator|O operador a ser usado para validar os dados.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriedade_ > enableMultipleFilterItems|Determina se deseja permitir vários itens de filtro.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriedade_ > id|ID do FilterPivotHierarchy. Somente leitura.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriedade_ > nome|Nome do FilterPivotHierarchy.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriedade_ > posição|Posição do FilterPivotHierarchy.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Relação_ > fields|Retorna PivotFields associados a FilterPivotHierarchy. Somente leitura.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Método_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|Redefina a FilterPivotHierarchy para os valores padrão.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Propriedade_ > itens|Um conjunto de objetos filterPivotHierarchy. Somente leitura.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual. Se houver hierarquia em outro lugar na linha, coluna ou eixo de filtro, ele será removido desse local.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém FilterPivotHierarchy por nome ou id.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém um FilterPivotHierarchy por nome. Se o FilterPivotHierarchy não existir, retornará um objeto nulo.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1,8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Propriedade_ > inCellDropDown|Exibido na lista na célula suspensa ou não, ele será padronizado como verdadeiro.|1,8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Propriedade_ > source|Fonte da lista de validação de dados|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriedade_ > id|ID do PivotField.. Somente leitura.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriedade_ > nome|Nome do PivotField.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriedade_ > showAllItems|Determina se deseja mostrar todos os itens de PivotField.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relação_ > items|Retorna PivotFields associados ao PivotField. Somente leitura.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relação_ > subtotals|Subtotais de PivotField.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Método_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|Classifica o PivotField. Se um DataPivotHierarchy for especificado, a classificação será aplicada com base nele, se a classificação não for baseada no campo PivotField.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Propriedade_ > itens|Um conjunto de objetos pivotField. Somente leitura.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Método_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Obtém PivotHierarchy por nome ou id.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Obtém o PivotHierarchy por nome. Se o PivotHierarchy não existir, retornará um objeto null.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Propriedade_ > id|ID do PivotHierarchy. Somente leitura.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Propriedade_ > nome|Nome do PivotHierarchy.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Relação_ > fields|Retorna PivotFields associados a PivotHierarchy. Somente leitura.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Propriedade_ > itens|Um conjunto de objetos pivotHierarchy. Somente leitura.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Método_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Obtém PivotHierarchy por nome ou id.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Obtém o PivotHierarchy por nome. Se o PivotHierarchy não existir, retornará um objeto null.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriedade_ > id|ID do PivotItem. Somente leitura.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriedade_ > isExpanded|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriedade_ > nome|Nome do PivotItem.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriedade_ > visible|Determina se o PivotItem ficará visível ou não.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Propriedade_ > itens|Um conjunto de objetos pivotItem. Somente leitura.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Método_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Obtém PivotHierarchy por nome ou id.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Obtém o PivotHierarchy por nome. Se o PivotHierarchy não existir, retornará um objeto null.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriedade_ > showColumnGrandTotals|Verdadeiro, quando o relatório de Tabela Dinâmica mostra os totais de colunas.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriedade_ > showRowGrandTotals|Verdadeiro, quando o relatório de Tabela Dinâmica mostra os totais de linhas.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriedade_ > subtotalLocation|Essa propriedade indica SubtotalLocationType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo. Valores possíveis são: AtTop, AtBottom.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Relação_ > layoutType|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os rótulos de coluna da Tabela Dinâmica.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os valores de dados da tabela dinâmica.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo de área de filtro da Tabela Dinâmica.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo em que a Tabela Dinâmica existe, excluindo a área de filtro.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os rótulos de linha da Tabela Dinâmica.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > columnHierarchies|As hierarquias de pivô da coluna da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > dataHierarchies|As hierarquias dinâmicas de dados da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > filterHierarchies|As hierarquias de pivô do filtro da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > hierarchies|Hierarquias pivô da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > layout|O PivotLayout descreve o layout e estrutura visual da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > rowHierarchies|As hierarquias de pivô de linha da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Método_ > [delete()](/javascript/api/excel/excel.pivottable)|Exclui a Tabela Dinâmica.|1,8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > [adicionar (nome: cadeia de caracteres, fonte: objeto, destino: objeto)](/javascript/api/excel/excel.pivottablecollection)|Adiciona um Pivottable com base nos dados de origem especificados e insere-o na célula superior esquerda do intervalo de destino.|1,8|
|[range](/javascript/api/excel/excel.range)|_Relação_ > dataValidation|Retorna um objeto de validação de dados. Somente leitura.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriedade_ > id|ID do RowColumnPivotHierarchy. Somente leitura.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriedade_ > nome|Nome da RowColumnPivotHierarchy.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriedade_ > posição|Posição da RowColumnPivotHierarchy.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Relação_ > fields|Retorna PivotFields associados a RowColumnPivotHierarchy. Somente leitura.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Método_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|Redefine o RowColumnPivotHierarchy para os valores padrão.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Propriedade_ > itens|Um conjunto de objetos rowColumnPivotHierarchy. Somente leitura.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual. Se houver a hierarquia em outro lugar na linha, coluna,|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém RowColumnPivotHierarchy por nome ou id.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém um RowColumnPivotHierarchy por nome. Se o RowColumnPivotHierarchy não existir, retornará um objeto nulo.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1,8|
|[runtime](/javascript/api/excel/excel.runtime)|_Propriedade_ > enableEvents|Alterna os eventos JavaScript no painel de tarefas atual ou no suplemento de conteúdo.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relação_ > baseField|O PivotField base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relação_ > baseItem|O Item base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relação_ > calculation|O cálculo de ShowAs a ser usado para o Data PivotField.|1,8|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > autoIndent|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|1,8|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > textOrientation|A orientação de texto para o estilo.|1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > automatic|Se Automatic for definido como true, todos os outros valores serão ignorados ao definir os subtotais.|1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > average| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > count| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > countNumbers| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > max| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > min| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > product| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > standardDeviation| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > standardDeviationP| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > sum| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > variance| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > varianceP| |1,8|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > legacyId|Retorna uma identificação numérica. Somente leitura.|1,8|
|[workbook](/javascript/api/excel/excel.workbook)|_Propriedade_ > readOnly|True se a pasta de trabalho estiver aberta no modo somente leitura. Somente leitura.|1,8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Propriedade_ > id|Retorna um valor que identifica de forma exclusiva o objeto WorkbookCreated. Somente leitura.|1,8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Método_ > [Open()](/javascript/api/excel/excel.workbookcreated)|Abra a pasta de trabalho.|1,8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > showGridlines|Obtém ou define um sinalizador de linhas de grade da planilha.|1,8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > showHeadings|É ou define um sinalizador de cabeçalhos da planilha.|1,8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha que é calculada.|1,8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Quais são as novidades na API JavaScript do Excel 1.7

O conjunto de requisitos 1.7 da API JavaScript do Excel incluei APIs para gráficos, eventos, planilhas, intervalos, propriedades do documento, itens nomeados, opções de proteção e estilos.

### <a name="customize-charts"></a>Personalize gráficos

Com as novas APIs de gráficos, você pode criar tipos degráficos adicionais, adicionar uma série de dados a um gráfico, definir o título do gráfico, adicionar um título de eixo, adicionar unidade de exibição, adicionar uma linha de tendência com média móvel, alterar uma linha de tendência para linear e muito mais. Estes são alguns exemplos:

* Eixo gráfico - obtenha, defina, formate e remova unidade de eixo, etiqueta e título em um gráfico.
* Série de gráficos - adicione, defina e exclua uma série em um gráfico.  Alterar marcadores da série, pedidos de plotagem e dimensionamento.
* Gráfico de linhas de tendências: adicione, receba e formate linhas de tendências em um gráfico.
* Legenda do gráfico - formate a fonte de legenda de um gráfico.
* Ponto do gráfico - defina a cor do ponto do gráfico.
* Subtítulo do título do gráfico - obtenha e defina a subseqüência do título para um gráfico.
* Tipo de gráfico - opção para criar mais tipos de gráfico.

### <a name="events"></a>Eventos

As APIs de eventos JavaScript do Excel fornecem diversos,  manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Você pode criar essa função para executar as ações que seu cenário exige. Para obter uma lista de eventos que estão disponíveis, confira [trabalhar com eventos usando as API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar a aparência de planilhas e intervalos

Nas novas APIs você pode personalizar a aparência das planilhas de várias maneiras:

* Congele painéis para manter linhas ou colunas específicas visíveis durante a rolagem na planilha. Por exemplo, se a primeira linha da planilha inclui cabeçalhos, você pode congelá-la para que os cabeçalhos das colunas permaneçam visíveis enquanto rola para baixo na planilha.
* Modificar a cor da guia de planilha.
* Adicione títulos de planilha.


Você pode personalizar a aparência de intervalos de várias maneiras:

* Defina o estilo de célula para um intervalo para garantir que todas as células no intervalo tenham formatação consistente. Um estilo de célula é um conjunto definido de características de formatação, como fontes e tamanhos de fonte, formatos numéricos, bordas de célula e sombreamento de célula. Use qualquer um dos estilos de célula internas do Excel ou crie seu próprio estilo de célula personalizado.
* Defina a orientação de texto para um intervalo.
* Adicione ou modifique um hiperlink em um intervalo vinculado a outro local na pasta de trabalho ou a um local externo.

### <a name="manage-document-properties"></a>Gerenciar propriedades dos documentos

Usando as APIs de propriedades do documento, você pode acessar as propriedades do documento interno e também criar e gerenciar propriedades personalizadas do documento para armazenar o estado da pasta de trabalho e direcionar o fluxo de trabalho e a lógica comercial.

### <a name="copy-worksheets"></a>Copiar planilhas

Usando a cópia da planilha APIs, você pode copiar os dados e o formato de uma planilha para uma nova planilha na mesma pasta de trabalho e reduzir a quantidade de transferência de dados necessária.

### <a name="handle-ranges-with-ease"></a>Lidar com intervalos com facilidade

Usando várias APIs de intervalo, você pode fazer coisas como obter região ao redor, obter um intervalo redimensionado e muito mais. Essas APIs devem tornar as tarefas, como manipulação de intervalo e endereçamento, muito mais eficientes.

Além disso:

* Opções de proteção de pasta de trabalho e planilha - use estas APIs para proteger dados em uma planilha e a estrutura da pasta de trabalho.
* Atualizar um item nomeado - usar esta API para atualizar um item nomeado.
* Obter a célula ativa - usar esta API para acessar a célula ativa da pasta de trabalho.

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > chartType|Representa o tipo de gráfico. Valores ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > id|Id exclusiva do gráfico. Somente leitura.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > showAllFieldButtons|Representa se deseja exibir todos os botões de campo em um Gráfico Dinâmico.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Relação_ > border|Representa o formato da borda da área de gráfico, incluindo a cor, estilo de linha e espessura. Somente leitura.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Método_ > getItem (tipo: cadeia de caracteres, grupo: cadeia de caracteres)|Retorna o eixo específico identificado por tipo e grupo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > axisBetweenCategories|Representa se o eixo de valor cruza o eixo de categoria entre as categorias.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > axisGroup|Representa o grupo para o eixo especificado. Somente leitura. Os valores possíveis são: Primary, Secondary.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > categoryType|Retorna ou define o tipo de eixo de categoria. Os valores possíveis são: TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > crosses.|Representa eixo especificado onde o outro eixo cruza. Os valores possíveis são: Automatic, Maximum, Minimum, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > crossesAt|Representa eixo especificado onde o outro eixo cruza. Somente leitura. A definição para essa propriedade deve usar o método SetCrossesAt (duplo). Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > customDisplayUnit|Representa o valor da unidade de exibição do eixo personalizado. Somente leitura. Para definir essa propriedade, use o método de SetCustomDisplayUnit(duplo). Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > displayUnit|Representa a unidade de exibição de eixo. Os valores possíveis são: None, Hundreds, Thousands, TenThousands, HundredThousands, Millions, TenMillions, HundredMillions, Billions, Trillions, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > height|Representa a altura, em pontos, do eixo do gráfico. Nulo se o eixo não for visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > left|Representa a distância, em pontos, da borda esquerda do eixo à esquerda da área do gráfico. Nulo se o eixo não for visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > logBase|Representa a base do logaritmo ao usar escalas logarítmicas.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > reversePlotOrder|Representa se o Microsoft Excel plota os pontos de dados do último para o primeiro.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > scaleType|Representa o tipo de escala do eixo dos valores. Valores possíveis são: Linear, Logarithmic.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > showDisplayUnitLabel|Indica se a etiqueta de unidade de exibição de eixo está visível.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > tickLabelSpacing|Representa o número série ou categorias entre os rótulos de marcas de escala. Pode ser um valor de 1 a 31999 ou uma cadeia de caracteres vazia para configuração automática. O valor retornado sempre é um número.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > tickMarkSpacing|Representa o número de série ou categorias entre as marcas de escala.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > top|Representa a distância, em pontos, da borda superior do eixo a parte superior da área do gráfico. Nulo se o eixo não for visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > type|Representa o tipo de eixo. Somente leitura. Os valores possíveis são: Invalid, Category, Value, Series.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > visible|Um valor booliano representa a visibilidade do eixo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > width|Representa a largura, em pontos, do eixo do gráfico. Nulo se o eixo não for visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > baseTimeUnit|Retorna ou define a unidade base para o eixo da categoria especificada.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > majorTickMark|Representa o tipo de marca de escala principal para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > majorTimeUnitScale|Retorna ou define o valor de escala de unidades principais para o eixo das categorias quando a propriedade CategoryType estiver definida como escala de tempo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > minorTickMark|Representa o tipo de marca de escala secundária para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > minorTimeUnitScale|Retorna ou define o valor da escala unitária secundária para o eixo da categoria quando a propriedade CategoryType estiver definida como TimeScale.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > tickLabelPosition|Representa a posição dos rótulos de marcas de escala no eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCategoryNames(sourceData: Range)|Define todos os nomes de categoria para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCrossesAt(valor: duplo)|Define o eixo especificado onde o outro eixo cruza.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCustomDisplayUnit(valor: duplo)|Definirá a unidade de exibição de eixo a um valor personalizado.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propriedade_ > color|Código de cor HTML que representa a cor das bordas no gráfico.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propriedade_ > espessura|Representa a espessura da borda, em pontos.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Relação_ > lineStyle|Representa o estilo de linha da borda.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > posição|Valor de DataLabelPosition que representa a posição do rótulo de dados. Os valores possíveis são: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > separator|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showBubbleSize|Valor booliano que determina se o tamanho da bolha do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showCategoryName|Valor booliano que determina se o nome da categoria do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showLegendKey|Valor booliano que determina se o código de legenda do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showPercentage|Valor booliano que determina se o percentual do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showSeriesName|Valor booliano que determina se o nome da série do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showValue|Valor booliano que determina se o valor do rótulo de dados fica visível ou não.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > height|Representa a altura da legenda no gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > left|Representa a esquerda de uma legenda do gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > showShadow|Representa se a legenda tem sombra no gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > top|Representa o início de uma legenda do gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > width|Representa a largura da legenda no gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Relação_ > legendEntries|Representa uma coleção de legendEntries na legenda. Somente leitura.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > visible|Representa o visível de uma entrada de legenda do gráfico.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Propriedade_ > itens|Um conjunto de objetos chartLegendEntry. Somente leitura.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Método_ > getCount()|Retorna o número de legendEntry da coleção.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Método_ > getItemAt(index: número)|Retorna legendEntry no índice fornecido.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > hasDataLabel|Representa se um ponto de dados possui um datalabel. Não aplicável para gráficos de superfície.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerBackgroundColor|Representação do código de cor HTML da cor de fundo do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerForegroundColor|Representação do código de cor HTML da cor de primeiro plano do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerSize|Representa o tamanho do marcador do ponto de dados.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerStyle|Representa estilo do marcador de um ponto de dados do gráfico. Os valores possíveis são: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Relação_ > dataLabel|Retorna o rótulo de dados de um ponto de gráfico. Somente leitura.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Relação_ > border|Representa o formato da borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e a espessura. Somente leitura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > chartType|Representa o tipo de gráfico de uma série. Valores ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > doughnutHoleSize|Representa o tamanho do furo de rosca de uma série de gráficos.  Válida apenas em gráficos de rosca e doughnutExploded.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > filtered|Valor booliano representando se a série é filtrada ou não. Não aplicável para gráficos de superfície.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > gapWidth|Representa a largura do espaçamento de uma série de gráfico.  Válida apenas sobre gráficos de barras e colunas, bem como|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > hasDataLabels|Valor booliano representando se a série tem rótulos de dados ou não.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerBackgroundColor|Representa a cor de fundo dos marcadores de uma série de gráficos.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerForegroundColor|Representa cor de primeiro plano dos marcadores de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerSize|Representa o tamanho do marcador de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerStyle|Representa o estilo do marcador de uma série de gráfico. Os valores possíveis são: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > plotOrder|Representa a ordem de plotagem de uma série de gráficos dentro do grupo de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > showShadow|Valor booliano representando se a série tem sombra ou não.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > smooth|Valor booliano representando se a série é suave ou não. Apenas para gráficos de linha e de dispersão.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > dataLabels|Representa uma coleção de todos os dataLabels da série. Somente leitura.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > linhas de tendência|Representa uma coleção de todas as linha de tendência da série. Somente leitura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > Delete()|Exclui a série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setBubbleSizes(sourceData: Range)|Definir tamanhos das bolhas para uma série de gráfico. Funciona apenas para gráficos de bolhas.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setValues(sourceData: Range)|Definir valores de uma série de gráficos. Para gráfico de dispersão, isso significa valores do eixo Y.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setXAxisValues(sourceData: Range)|Definir valores do eixo X para uma série de gráficos. Funciona apenas para gráficos de dispersão.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Método_ > add (nome: cadeia de caracteres, indexar: número)|Adiciona uma nova série para o conjunto.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > height|Representa a altura, em pontos, do título do gráfico. Somente leitura. Nulo se o título do gráfico não estiver visível. Somente leitura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > horizontalAlignment|Representa o alinhamento horizontal para título do gráfico. Os valores possíveis são: Center, Left, Justify, Distributed, Right.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > left|Representa a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico. Nulo se o título do gráfico não estiver visível.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > posição|Representa a posição de título do gráfico. Os valores possíveis são: Top, Automatic, Bottom, Right, Left.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > showShadow|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > textOrientation|Representa a orientação de texto do título do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > top|Representa a distância em pontos, da borda superior do título do gráfico a parte superior da área do gráfico. Nulo se o título do gráfico não estiver visível.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > verticalAlignment|Representa o alinhamento vertical do título do gráfico. Os valores possíveis são: Center, Bottom, Top, Justify, Distributed.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > width|Retorna a largura em pontos do título do gráfico. Somente leitura. Nulo se o título do gráfico não estiver visível. Somente leitura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Método_ > setFormula(fórmula: cadeia de caracteres)|Define um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Relação_ > border|Representa o formato da borda do título do gráfico, incluindo a cor, estilo de linha e espessura. Somente leitura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > backward|Representa o número de períodos que a linha de tendência se estende para trás.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > displayEquation|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > displayRSquared|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > forward|Representa o número de períodos que a linha de tendência se estende para frente.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > intercept|Representa o valor de intercepção da linha de tendência. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo). O valor retornado sempre é um número.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > movingAveragePeriod|Representa o período de uma linha de tendência do gráfico, apenas para a linha de tendência com o tipo MovingAverage.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > nome|Representa o nome da linha de tendência. Pode ser definido como um valor de sequência ou pode ser definido como valor nulo para representar valores automáticos. O valor retornado sempre é uma cadeia de caracteres.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > polynomialOrder|Representa a ordem de uma linha de tendência do gráfico, apenas para a linha de tendência com o tipo Polynomial.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > type|Representa o tipo da linha de tendência de um gráfico. Valores possíveis são: Linear, Exponential, Logarithmic, MovingAverage, Polynomial, Power.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relação_ > format|Representa a formatação de uma linha de tendência do gráfico. Somente leitura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Método_ > Delete()|Deleta o objeto Trendline.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Propriedade_ > itens|Um conjunto de objetos chartTrendline. Somente leitura.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > add(type: string)|Adiciona uma nova linha de tendência ao conjunto de linha de tendência.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > getCount()|Retorna o número de linha de tendência na coleção.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > getItem(index: number)|Obtém o objeto da linha de tendência por índice, que é a ordem de inserção na matriz de itens.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Relação_ > line|Representa a formatação de linha do gráfico. Somente leitura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > key|Obtém a chave da propriedade personalizada. Somente leitura. Somente leitura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > type|Obtém o tipo de valor da propriedade personalizada. Somente leitura. Somente leitura. Os valores possíveis são: Number, Boolean, Date, String, Float.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > value|Obtém ou define o valor da propriedade personalizada.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Método_ > Delete()|Exclui a propriedade personalizada.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Propriedade_ > itens|Uma coleção de objetos customProperty. Somente leitura.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > add (key: string, value: object)|Cria uma nova propriedade personalizada ou define uma existente.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > deleteAll()|Exclui todas as propriedades personalizadas nesta coleção.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getCount()|Obtém a contagem das propriedades personalizadas.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getItem(key: string)|Obtém um objeto de propriedade personalizado por sua chave, que não faz distinção entre maiúsculas e minúsculas. Lança se a propriedade customizada não existir.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getItemOrNullObject(key: string)|Obtém um objeto de propriedade personalizado por sua chave, que não faz distinção entre maiúsculas e minúsculas. Retorna um objeto nulo se a propriedade customizada não existir..|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Propriedade_ > itens|Um conjunto de objetos de conexão de dados. Somente leitura.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Método_ > refreshAll()|Atualiza todas as conexões de dados da coleção.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > author|Obtém ou define o autor da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > category|Obtém ou define a categoria da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > comments|Obtém ou define os comentários da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > company|Obtém ou define a empresa do documento.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > keywords|Obtém ou define as palavras-chave da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > lastAuthor|Obtém o último autor da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > manager|Obtém ou define o gerenciador da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > revisionNumber|Obtém o número de revisão da pasta de trabalho. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > subject|Obtém ou define o assunto da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > title|Obtém ou define o título da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relação_ > creationDate|Obtém a data de criação da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relação_ > custom|Obtém a coleção de propriedades personalizadas da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > formula|Obtém ou define a fórmula do item nomeado.  A fórmula sempre começa com um sinal de "=".|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relação_ > arrayValues|Retorna um objeto que contém valores e tipos do item nomeado. Somente leitura.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propriedade_ > types|Representa os tipos de cada item na matriz de itens nomeados como somente leitura. Os valores possíveis são: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propriedade_ > values|Representa os valores de cada item na matriz de itens nomeados. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > isEntireColumn|Representa se o intervalo atual está em uma coluna inteira. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > isEntireRow|Representa se o intervalo atual está em uma linha inteira. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > numberFormatLocal|Representa o código de formato numérico do Excel para o intervalo fornecido como uma cadeia de caracteres no idioma do usuário.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > style|Representa o estilo de intervalo atual. Isso retornará nulo ou uma cadeia de caracteres.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getAbsoluteResizedRange (numRows: número numColumns: número)|Obtém um objeto Range com a mesma célula superior esquerda do objeto Range atual, mas com os números especificados de linhas e colunas.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getImage()|O intervalo é renderizado como uma imagem em base 64.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getSurroundingRegion()|Retorna um objeto Range que representa a região circundante da célula superior esquerda nesse intervalo. Uma região ao redor é um intervalo limitado por qualquer combinação de linhas e colunas em branco em relação a esse intervalo.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > showCard()|Exibe o cartão para uma célula ativa se ele tiver um conteúdo valioso.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > textOrientation|Obtém ou define a orientação de texto de todas as células no intervalo.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > useStandardHeight|Determina se a altura da linha do objeto Range é igual a altura padrão da planilha.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > useStandardWidth|Determina se a largura da coluna do objeto Range é igual a largura padrão da planilha.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > address|Representa o destino da url do hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > document.|Representa o documento. meta do hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > screenTip|Representa a cadeia exibida ao passar o mouse sobre o hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > textToDisplay|Representa a cadeia de caracteres exibida na parte superior esquerda da maioria das células no intervalo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > addIndent|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > autoIndent|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > builtIn|Indica se o estilo é um estilo interno. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > formulaHidden|Indica se a fórmula ficará oculta quando a planilha estiver protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > horizontalAlignment|Representa o alinhamento horizontal para o estilo. Os valores possíveis são: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeAlignment|Indica se o estilo incluem as propriedades AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, e TextOrientation.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeBorder|Indica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeFont|Indica se o estilo inclui as propriedades de fonte Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript e Underline.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeNumber|Indica se o estilo inclui a propriedade NumberFormat.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includePatterns|Indica se o estilo inclui as propriedades internas Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeProtection|Indica se o estilo incluirá as propriedades de proteção FormulaHidden e Locked.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > indentLevel|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > locked|Indica se o objeto é bloqueado quando a planilha está protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > nome|O nome do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > numberFormat|O código de formatação de formato de número para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > numberFormatLocal|O código de formato localizado do formato numérico para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > orientation|A orientação de texto para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > readingOrder|A ordem de leitura para o estilo. Os valores possíveis são: Context, LeftToRight, RightToLeft.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > shrinkToFit|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > textOrientation|A orientação de texto para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > verticalAlignment|Representa o alinhamento vertical do estilo. Os valores possíveis são: Top, Center, Bottom, Justify, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > wrapText|Indica se o Microsoft Excel quebra automaticamente a linha de texto no objeto.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relação_ > borders|Uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relação_ > fill|O preenchimento do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relação_ > font|Objeto de fonte que representa a fonte do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Método_ > Delete()|Exclui este estilo.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Propriedade_ > itens|Uma coleção de objetos de estilo. Somente leitura.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Método_ > add(name: string)]|Adiciona um novo estilo para o conjunto.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Método_ > getItem(name: string)|Obtém um estilo por nome.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > address|Obtém o endereço que representa a área alterada de uma tabela em uma planilha específica.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > changeType|Obtém o tipo de mudança que representa como o evento Changed é acionado. Os valores possíveis são: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > tableId|Obtém o id da tabela na qual os dados foram alterados.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual os dados são alterados.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > address|Obtém o endereço do intervalo que representa a área selecionada da tabela em uma planilha específica.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > isInsideTable|Indica se a seleção está dentro de uma tabela, o endereço será inútil se IsInsideTable for falso.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > tableId|Obtém o id da tabela na qual a seleção foi alterada.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual a seleção foi alterada.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Propriedade_ > nome|Obtém o nome da pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > dataConnections|Atualiza todas as conexões de dados na pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > properties|Obtém as propriedades da pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > protection|Retorna o objeto de proteção de pasta de trabalho para uma pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > styles|Representa uma coleção de estilos associados à pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Método_ > getActiveCell()|Obtém a célula ativa no momento da pasta de trabalho.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Propriedade_ > protected|Indica se a pasta de trabalho está protegida. Somente Leitura. Somente leitura.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Método_ > protect(password: string)|Protege uma pasta de trabalho. Falhará se a pasta de trabalho estiver protegida.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Método_ > unprotect(password: string)|Desprotege uma pasta de trabalho.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > gridlines|Obtém ou define um sinalizador de linhas de grade da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > headings|É ou define um sinalizador de cabeçalhos da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > showHeadings|É ou define um sinalizador de cabeçalhos da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > standardHeight|Retorna a altura padrão de todas as linhas na planilha, em pontos. Somente leitura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > standardWidth|Retorna ou define a largura padrão de todas as colunas na planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > tabColor|Obtém ou define a cor da guia de planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > freezePanes|Obtém um objeto que pode ser usado para manipular painéis congelados na planilha somente leitura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|Copia uma planilha e a coloca na posição especificada. Retorna à planilha copiada.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha que está ativada.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha que é adicionada à pasta de trabalho.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > address|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > changeType|Obtém o tipo de mudança que representa como o evento Changed é acionado. Os valores possíveis são: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual os dados são alterados.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha que está desativada.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > worksheetId|Obtém o id do gráfico que é excluído da pasta de trabalho.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeAt(frozenRange: Range or string)|Define as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeColumns(count: number)|Congela a primeira colunas da planilha no local.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeRows(count: number)|Congela as linhas superiores da planilha no local.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > getLocation()|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > getLocationOrNullObject()|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > unfreeze()|Remove todos os painéis congelados na planilha.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowEditObjects|Indica a opção de proteção de planilha para permitir a edição de objetos.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowEditScenarios|Indica a opção de proteção de planilha para permitir a edição de cenários.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Relação_ > selectionMode|Representa a opção de proteção da planilha do modo de seleção.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > address|Obtém o endereço do intervalo que representa a área selecionada de uma planilha específica.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual a seleção foi alterada.|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Quais são as novidades na API JavaScript do Excel 1.6 

### <a name="conditional-formatting"></a>Formatação condicional

Introduz a formatação condicional de um intervalo. Permite os seguintes tipos de formatação condicional:

* Escala de cores
* Barra de dados
* Conjunto de ícones
* Personalizado

Além disso:

* Retorna o intervalo ao qual o formatato condicional é aplicada. 
* Remoção da formatação condicional. 
* Fornece a capacidade de priority e stopifTrue. 
* Obtém a coleção de toda a formatação condicional em um determinado intervalo. 
* Limpa todos os formatos condicionais ativos no intervalo atual especificado. 

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Método_ > suspendApiCalculationUntilNextSync()|Suspende o cálculo até que o próximo "context.sync()" seja chamado. Uma vez definido, é responsabilidade do desenvolvedor recalcular a pasta de trabalho, para garantir que todas as dependências sejam propagadas.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relação_ > rule|Representa o objeto Regra neste formato condicional.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Propriedade_ > threeColorScale|Caso verdadeiro, a escala de cores terá três pontos (mínimo, médio, máximo). Caso contrário, terá dois (mínimo, máximo). Somente leitura.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Relação_ > criteria|Os critérios da escala de cores. O ponto médio é opcional ao se usar uma escala de cores de dois pontos.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > formula1|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > formula2|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > operator|O operador do formato condicional de texto. Os valores possíveis são: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relação_ > maximum|O critério de escala de cores de ponto máximo.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relação_ > midpoint|O critério de escala de cores de ponto médio, se a escala de cores for uma escala de três cores.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relação_ > minimum|O critério de escala de cores de ponto mínimo.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > color|Representação de código de cor HTML da cor de escala de cores. Por exemplo, #FF0000 representa vermelho.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > formula|Um número, uma fórmula ou nulo (se Type for LowestValue).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > type|No que a fórmula condicional de ícone deve se basear. Os valores possíveis são: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > borderColor|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > fillColor|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > matchPositiveBorderColor|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > matchPositiveFillColor|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > borderColor|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > fillColor|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > gradientFill|Representação booliana para indicar se a DataBar tem um gradiente ou não.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propriedade_ > formula|A fórmula, se necessário, para avaliar a regra databar.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propriedade_ > type|O tipo de regra para databar. Os valores possíveis são: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > id|A prioridade do formato condicional na atual ConditionalFormatCollection. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > priority|A prioridade (ou índice) dentro da coleção de formatos condicionais na qual se encontra atualmente esse formato condicional. Alterando isso também.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > stopIfTrue|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > type|Um tipo de formatação condicional. É possível definir somente um por vez. Somente leitura. Os valores possíveis são: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > cellValue|Retornará as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo de CellValue. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > cellValueOrNullObject|Retornará as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo de CellValue. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > colorScale|Retornará as propriedades de formato condicional de ColorScale se o formato condicional atual for um tipo de ColorScale. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > colorScaleOrNullObject|Retornará as propriedades de formato condicional de ColorScale se o formato condicional atual for um tipo de ColorScale. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > custom|Retornará as propriedades personalizadas do formato condicional se o formato condicional atual for um tipo personalizado. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > customOrNullObject|Retornará as propriedades personalizadas do formato condicional se o formato condicional atual for um tipo personalizado. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > dataBar|Retornará as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > dataBarOrNullObject|Retornará as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > iconSet|Retornará as propriedades do formato condicional de IconSet se o formato condicional atual for um tipo de IconSet. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > iconSetOrNullObject|Retornará as propriedades do formato condicional de IconSet se o formato condicional atual for um tipo de IconSet. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > preset|Retornará o formato condicional de critérios predefinidos, como as propriedades above averagebelow averageunique valuescontains blanknonblankerrornoerror.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > presetOrNullObject|Retornará o formato condicional de critérios predefinidos, como as propriedades above averagebelow averageunique valuescontains blanknonblankerrornoerror.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > textComparison|Retornará as propriedades específicas do formato condicional de texto se o formato condicional atual for um tipo de texto.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > textComparisonOrNullObject|Retornará as propriedades específicas do formato condicional de texto se o formato condicional atual for um tipo de texto.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > topBottom|Retornará as propriedades do formato condicional de TopBottom se o formato condicional atual for um tipo de TopBottom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > topBottomOrNullObject|Retornará as propriedades do formato condicional de TopBottom se o formato condicional atual for um tipo de TopBottom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > Delete()|Exclui esse formato condicional.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > getRange()|Retornará o intervalo ao qual o formato condicional está aplicado ou um objeto nulo se o intervalo for descontínuo. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > getRangeOrNullObject()|Retornará o intervalo ao qual o formato condicional está aplicado ou um objeto nulo se o intervalo for descontínuo. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Propriedade_ > itens|Uma coleção de objetos conditionalFormat. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > add(type: string)|Adiciona um novo formato condicional à coleção na prioridade firsttop.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > clearAll()|Limpa todos os formatos condicionais ativos no intervalo atual especificado.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getCount()|Retorna o número de formatos condicionais na pasta de trabalho. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getItem(id: string)|Retorna um formato condicional para o ID fornecido.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getItemAt(index: número)|Retorna um formato condicional no índice fornecido.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > formula|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > formulaLocal|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > formulaR1C1|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propriedade_ > formula|Um número ou uma fórmula, dependendo do tipo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propriedade_ > operator|GreaterThan ou GreaterThanOrEqual para cada tipo de regra para o formato de ícone condicional. Os valores possíveis são Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relação_ > customIcon|O ícone personalizado para o critério atual, se diferente do IconSet padrão; caso contrário, será retornado nulo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relação_ > type|No que a fórmula condicional de ícone deve se basear.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Propriedade_ > criterion|O critério do formato condicional. Os valores possíveis são: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > color|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > id|Representa o identificador da borda. Somente leitura. Os valores possíveis são: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > sideIndex|Valor constante que indica o lado específico da borda. Somente leitura. Os valores possíveis são: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > style|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Os valores possíveis são: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propriedade_ > count|Número de objetos de borda da coleção. Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propriedade_ > itens|Uma coleção de objetos conditionalRangeBorder. Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relação_ > bottom|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relação_ > left|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relação_ > right|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relação_ > top|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Método_ > getItem(index: string)|Obtém um objeto de borda usando seu nome|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Método_ > getItemAt(index: número)|Obtém um objeto de borda usando seu índice.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Propriedade_ > color|Código de cor HTML que representa a cor do preenchimento do formulário #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Método_ > clear()|Redefine o preenchimento.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > bold|Representa o status da fonte em negrito.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > color|Representação de código de cor HTML para a cor do texto. Por exemplo, #FF0000 representa vermelho.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > italic|Representa o status da fonte em itálico.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > strikethrough|Representa o status de tachado da fonte.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > underline|Tipo de sublinhado aplicado à fonte. Os valores possíveis são: None, Single, Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Método_ > clear()|Redefine os formatos de fonte.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Propriedade_ > numberFormat|Representa o código de formato numérico do Excel para determinado intervalo. Desmarcado se o nulo for passado.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relação_ > borders|Coleção de objetos de borda que se aplicam ao intervalo de formatos condicionais geral. Somente leitura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relação_ > fill|Retorna o objeto de preenchimento definido no intervalo de formatos condicionais gerais. Somente leitura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relação_ > font|Retorna o objeto de fonte definido no intervalo de formatos condicionais gerais. Somente leitura.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propriedade_ > operator|O operador do formato condicional de texto. Os valores possíveis são: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propriedade_ > text|O valor de texto do formato condicional.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propriedade_ > rank|A classificação entre 1 e 1000 para classificações numéricas ou 1 e 100 para classificações percentuais.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propriedade_ > type|Formatar valores com base na classificação superior ou inferior. Os valores possíveis são: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relação_ > rule|Representa o objeto Regra neste formato condicional. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > axisColor|Código de cor HTML que representa a cor da linha de Eixo, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > axisFormat|Representação de como o eixo é determinado para uma barra de dados do Excel. Os valores possíveis são: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > barDirection|Representa a direção em que o gráfico de barras de dados deve ser baseado. Os valores possíveis são: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > showDataBarOnly|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relação_ > lowerBoundRule|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relação_ > negativeFormat|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relação_ > positiveFormat|Representação de todos os valores à direita do eixo em uma barra de dados do Excel. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relação_ > upperBoundRule|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > reverseIconOrder|Caso verdadeiro, inverte as ordens de ícones para IconSet. Observe que não será possível definir isso se ícones personalizados forem usados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > showIconOnly|Caso verdadeiro, oculta os valores e mostra somente ícones.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > style|Caso definido, exibe a opção IconSet do formato condicional. Os valores possíveis são: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Relação_ > criteria|Uma matriz de IconSets e critérios para as regras e os ícones personalizados potenciais para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto tipo, fórmula e operador serão ignorados quando definidos.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relação_ > rule|A regra da formatação condicional.|1.6|
|[range](/javascript/api/excel/excel.range)|_Relação_ > conditionalFormats|Coleção de ConditionalFormats que formam uma interseção do intervalo. Somente leitura.|1.6|
|[range](/javascript/api/excel/excel.range)|_Método_ > calculate()|Calcula um intervalo de células em uma planilha.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relação_ > rule|A regra da formatação condicional.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relação_ > rule|Os critérios da formatação condicional TopBottom.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > internalTest|Somente para uso interno. Somente leitura.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > calculate(markAllDirty: bool)|Calcula todas as células em uma planilha.|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Quais são as novidades na API JavaScript do Excel 1.5

### <a name="custom-xml-part"></a>Parte XML personalizada

* Adição de uma coleção de partes XML personalizadas ao objeto workbook.
* Obter parte XML personalizada usando ID
* Obtenção de um novo conjunto com escopo de partes XML personalizadas cujos namespaces correspondam ao namespace especificado.
* Obtenha uma cadeia XML associada a uma parte.
* Forneça id e namespace de uma parte.
* Adiciona uma nova parte XML personalizada à pasta de trabalho.
* Defina a parte XML inteira.
* Exclua uma parte XML personalizada.
* Exclua um atributo com o nome especificado do elemento identificado por xpath.
* Consulte o conteúdo XML por xpath.
* Insira, atualize e exclua o atributo.

**Implementação de referência:** Consulte [aqui](https://github.com/mandren/Excel-CustomXMLPart-Demo) para conhecer uma implementação de referência que mostra como partes XML personalizadas podem ser usadas em um suplemento.

### <a name="others"></a>Outros
* `range.getSurroundingRegion()` Retorna um objeto Range que representa a região ao redor desse intervalo. Uma região ao redor é um intervalo limitado por qualquer combinação de linhas e colunas em branco em relação a esse intervalo.
* `getNextColumn()` e `getPreviousColumn()`, `getLast() na coluna da tabela.
* `getActiveWorksheet()` na pasta de trabalho.
* `getRange(address: string)` fora da pasta de trabalho.
* `getBoundingRange(ranges: )` Obtém o menor objeto Range que abrange os intervalos fornecidos. Por exemplo, o intervalo delimitador entre "B2:C5" e "D10:E15" é "B2:E15".
* `getCount()` em várias coleções, como itens nomeados, planilhas, tabelas etc. para obter o número de itens em uma coleção. `workbook.worksheets.getCount()`
* `getFirst()` e `getLast()` e obter o último em várias coleções, como coleções de planilhas, colunas de tabela, pontos de gráfico e exibições de intervalo.
* `getNext()` e `getPrevious()` na coleção de planilhas e colunas de tabela.
* `getRangeR1C1()` Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Propriedade_ > id|ID da parte XML personalizada. Somente leitura.|1,5|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Propriedade_ > namespaceUri|URI do namespace da parte XML personalizada. Somente leitura.|1,5|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > Delete()|Exclui a parte XML personalizada.|1,5|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > getXml()|Obtém o conteúdo XML completo da parte XML personalizada.|1,5|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > setXml(xml: string)|Define o conteúdo XML completo da parte XML personalizada.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Propriedade_ > itens|Uma coleção de objetos customXmlPart. Somente leitura.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > add(xml: string)|Adiciona uma nova parte XML personalizada à pasta de trabalho.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getByNamespace(namespaceUri: string)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getCount()|Obtém o número de partes CustomXml na coleção.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getItem(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Propriedade_ > itens|Uma coleção de objetos customXmlPartScoped. Somente leitura.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getCount()|Obtém o número de partes CustomXML nesta coleção.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getItem(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getOnlyItem()|Se o conjunto contiver exatamente um item, esse método o retornará.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getOnlyItemOrNullObject()|Se o conjunto contiver exatamente um item, esse método o retornará.|1,5|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > customXmlParts|Representa a coleção de partes XML contidas nesta pasta de trabalho. Somente leitura.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getNext(visibleOnly: bool)|Obtém a planilha posterior a esta. Se não houver nenhuma planilha após esta, este método gerará um erro.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getNextOrNullObject(visibleOnly: bool)|Obtém a planilha posterior a esta. Se não houver nenhuma planilha após esta, este método retornará um objeto nulo.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getPrevious(visibleOnly: bool)|Obtém a planilha anterior a esta. Se não houver nenhuma planilha anterior, esse método lançará um erro.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getPreviousOrNullObject(visibleOnly: bool)|Obtém a planilha anterior a esta. Se não houver nenhuma planilha anterior, este método retornará um objeto nulo.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getFirst(visibleOnly: bool)|Obtém a primeira planilha na coleção.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getLast(visibleOnly: bool)|Obtém a última planilha na coleção.|1,5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Quais são as novidades na API JavaScript do Excel 1.4
A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.4.

### <a name="named-item-add-and-new-properties"></a>Adicionar item nomeado e novas propriedades

Novas propriedades:

* `comment`
* `scope` itens com escopo de planilha ou pasta de trabalho
* `worksheet` retorna a planilha que o item nomeado tem como escopo.

Novos métodos:

* `add(name: string, reference: Range or string, comment: string)`Adiciona um novo nome à coleção do escopo fornecido.
* `addFormulaLocal(name: string, formula: string, comment: string)` Adiciona um novo nome à coleção do escopo fornecido usando a localidade do usuário para a fórmula.

### <a name="settings-api-in-the-excel-namespace"></a>Configurações de API no namespace do Excel

O objeto [Configuração](/javascript/api/excel/excel.setting) representa um par chave-valor de uma configuração persistente ao documento. O recurso `Excel.Setting` é equivalente a `Office.Settings`, mas usa a sintaxe da API em lote, em vez de modelo de retorno de chamada de API comuns.

As APIs incluem `getItem()` para acessar configuração de entrada por meio da chave, `add()` para adicionar o par de configuração de chave:valor especificado na pasta de trabalho.

### <a name="others"></a>Outros

* Definir nome de coluna de tabela (a versão anterior permite somente leitura).
* Adicionar coluna de tabela ao fim da tabela (a versão anterior permite apenas em qualquer lugar, exceto o último).
* Adicione várias linhas a uma tabela de cada vez (a versão anterior só permite uma linha por vez).
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* Obter item ou função de objeto null: Esta funcionalidade permite obter o objeto utilizando a chave. Se o objeto não existir, a propriedade isNullObject do objeto retornado será true. Isso permite que os desenvolvedores verifiquem se existe um objeto ou não sem ter de lidar com ele por meio do tratamento de exceção. Disponível na planilha, item nomeado, associação, série de gráficos etc.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getCount()|Obtém o número de associações da coleção.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém um objeto binding pela ID. Se o objeto binding não existir, retornará um objeto null.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getCount()|Retorna o número de gráficos da planilha.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getItemOrNullObject(name: string)|Obtém um gráfico usando o respectivo nome. Quando houver vários gráficos com o mesmo nome, o sistema retornará o primeiro deles.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Método_ > getCount()|Retorna o número de pontos do gráfico da série.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Método_ > getCount()|Retorna o número de série da coleção.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > comment|Representa o comentário associado a esse nome.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > escopo|Indica se o nome tem escopo para a pasta de trabalho ou uma planilha específica. Somente leitura. Os valores possíveis são: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relação_ > planilha|Retorna a planilha em que o item nomeado tem escopo. Gerará um erro se os itens tiverem escopo para a pasta de trabalho em vez disso. Somente leitura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relação_ > worksheetOrNullObject|Retorna a planilha em que o item nomeado tem escopo. Retornará um objeto null se o item tiver escopo para a pasta de trabalho em vez disso. Somente leitura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Método_ > Delete()|Exclui o nome fornecido.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Método_ > getRangeOrNullObject()|Retorna o objeto Range associado ao nome. Retornará um objeto null se o tipo do item nomeado não for um intervalo.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > add(name: string, reference: Range or string, comment: string)|Adiciona um novo nome à coleção do escopo fornecido.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > addFormulaLocal (name: string, formula: string, comment: string)|Adiciona um novo nome à coleção de escopo fornecido usando a localidade do usuário para a fórmula.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getCount()|Obtém o número de itens nomeados na coleção.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getItemOrNullObject(name: string)|Obtém um objeto NamedItem usando o respectivo nome. Se o objeto getNamedItem não existir, retornará um objeto null.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getCount()|Obtém o número de tabelas dinâmicas na coleção.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItemOrNullObject(name: string)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, retornará um objeto null.|1.4|
|[range](/javascript/api/excel/excel.range)|_Método_ > getIntersectionOrNullObject(anotherRange: Intervalo ou cadeia de caracteres)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados. Se nenhuma interseção for encontrada, retornará um objeto null.|1.4|
|[range](/javascript/api/excel/excel.range)|_Método_ > getUsedRangeOrNullObject(valuesOnly: bool)|Retorna o intervalo usado do objeto range determinado. Se não houver nenhuma célula usada no intervalo, esta função retornará um objeto null.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Método_ > getCount()|Obtém o número de objetos RangeView na coleção.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > key|Retorna a chave que representa a id da configuração. Somente leitura.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > value|Representa o valor armazenado para esta configuração.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Método_ > Delete()|Exclui a configuração.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propriedade_ > itens|Uma coleção de objetos de configuração. Somente leitura.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > add(key: string, value: (any))|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getCount()|Obtém o número de Configurações na coleção.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItem(key: string)|Obtém uma entrada de configuração por meio da tecla.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItemOrNullObject(key: string)|Obtém uma entrada de configuração por meio da tecla. Se a Configuração não existir, retornará um objeto null.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relação_ > settings|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getCount()]|Obtém o número de tabelas na coleção.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getItemOrNullObject(key: number or string)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, retornará um objeto null.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getCount()|Obtém a quantidade de colunas na tabela.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getItemOrNullObject(key: number or string)|Obtém um objeto column por nome ou ID. Se a coluna não existir, retornará um objeto null.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Método_ > getCount()|Obtém a quantidade de linhas na tabela.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > settings|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > nomes|Coleção de nomes com escopo para a planilha atual. Somente leitura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getUsedRangeOrNullObject(valuesOnly: bool)|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, esta função retornará um objeto null.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getCount(visibleOnly: bool)|Obtém o número de planilhas na coleção.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getItemOrNullObject(key: string)|Obtém um objeto worksheet usando o Nome ou ID dele. Se a planilha não existir, retornará um objeto null.|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Quais são as novidades na API JavaScript do Excel 1.3

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.3.

|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Método_ > Delete()|Especifica a associação.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > add(range: Range or string, bindingType: string, id: string)|Adiciona uma nova associação a um intervalo específico.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromNamedItem(name: string, bindingType: string, id: string)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromSelection(bindingType: string, id: string)|Adiciona uma nova associação com base na seleção atual.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getItemOrNull(id: string)|Obtém um objeto de associação pela ID. Se o objeto em associação não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getItemOrNull(name: string)|Obtém um gráfico usando o respectivo nome. Quando houver vários gráficos com o mesmo nome, o sistema retornará o primeiro deles.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getItemOrNull(name: string)|Obtém um objeto NamedItem usando o respectivo nome. Se o objeto nameditem não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Propriedade_ > nome|Nome da Tabela Dinâmica.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > planilha|A planilha que contém a Tabela Dinâmica atual. Somente leitura.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Método_ > refresh()|Atualiza a Tabela Dinâmica.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Propriedade_ > itens|Uma coleção de objetos de Tabela Dinâmica. Somente leitura.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItem(name: string)|Obtém uma Tabela Dinâmica por nome.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItemOrNull(name: string)|Obtém uma Tabela Dinâmica por nome. Se a Tabela Dinâmica não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[range](/javascript/api/excel/excel.range)|_Método_ > getIntersectionOrNull(anotherRange: Range or string)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados. Se nenhuma interseção for encontrada, retornará um objeto null.|1.3|
|[range](/javascript/api/excel/excel.range)|_Método_ > getVisibleView()|Representa as linhas visíveis do intervalo atual.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > cellAddresses|Representa os endereços de célula da RangeView. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > columnCount|Retorna o número de colunas visíveis. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulas|Representa a fórmula em notação A1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulasLocal|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.  Por exemplo, a fórmula "=SUM(A1, introduced in 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulasR1C1|Representa a fórmula em notação no estilo L1C1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > index|Retorna um valor que representa o índice da RangeView. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > numberFormat|Representa o código de formato de número do Excel para determinada célula.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > rowCount|Retorna o número de linhas visíveis. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > texto|Valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > valueTypes|Representa o tipo de dados de cada célula. Somente leitura. Os valores possíveis são: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > values|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Relação_ > rows|Representa uma coleção de exibições de tabelas associadas ao intervalo. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Método_ > getRange()|Obtém o intervalo pai associado à RangeView atual.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Propriedade_ > itens|Uma coleção de objetos rangeView. Somente leitura.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Método_ > getItemAt(index: número)|Obtém uma linha de RangeView através de seu índice. Indexado com zero.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > key|Retorna a chave que representa a id da configuração. Somente leitura.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Método_ > Delete()|Exclui a configuração.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propriedade_ > itens|Uma coleção de objetos de configuração. Somente leitura.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItem(key: string)|Obtém uma entrada de configuração por meio da tecla.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItemOrNull(key: string)|Obtém uma entrada de configuração por meio da tecla. Se o objeto de configuração não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > set(key: string, value: string)|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relação_ > settingCollection|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > highlightFirstColumn|Indica se a primeira coluna contém uma formatação especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > highlightLastColumn|Indica se a última coluna contém uma formatação especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showBandedColumns|Indica se as colunas mostram formatação em faixas nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showBandedRows|Indica se as linhas mostram formatação em faixas nas quais as linhas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showFilterButton|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getItemOrNull(key: number or string)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getItemOrNull(key: number or string)|Obtém um objeto de coluna por nome ou ID. Se a coluna não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > pivotTables|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > settings|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > pivotTables|Coleção de Tabelas Dinâmicas que fazem parte da planilha. Somente leitura.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Quais são as novidades na API JavaScript do Excel 1.2

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.2.

|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > id|Obtém um gráfico com base em sua posição na coleção. Somente leitura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > planilha|A planilha que contém o gráfico atual. Somente leitura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Método_ > getImage(height: number, width: number, fittingMode: string)|Processa o gráfico como uma imagem codificada em base64, dimensionando o gráfico para se ajustar às dimensões especificadas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Relação_ > criteria|O filtro aplicado no momento à coluna fornecida. Somente leitura.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > apply(criteria: FilterCriteria)|Aplica os critérios de filtro determinados à coluna fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyBottomItemsFilter(count: number)|Aplica um filtro "Item Inferior" à coluna para obter o número de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyBottomPercentFilter(percent: number)]|Aplica um filtro "Percentual Inferior" à coluna para obter a porcentagem de elementos fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyCellColorFilter(color: string)|Aplica um filtro "Cor da Célula" à coluna para obter a cor fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyCustomFilter (criteria1: string, criteria2: string, oper: string)|Aplica um filtro "Ícone" à coluna para obter as cadeias de caracteres de critérios fornecidas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyDynamicFilter(criteria: string)|Aplica um filtro "Dinâmico" à coluna.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyFontColorFilter(color: string)|Aplica um filtro "Cor da Fonte" à coluna para obter a cor fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyIconFilter(icon: Icon)|Aplica um filtro "Ícone" à coluna para obter o ícone fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyTopItemsFilter(count: number)|Aplica um filtro "Item Superior" à coluna para obter o número de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyTopPercentFilter(percent: number)|Aplica um filtro "Percentual Superior" à coluna para obter a porcentagem de elementos fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyValuesFilter (valores: ())|Aplica um filtro "Valores" à coluna para obter os valores fornecidos.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > clear()|Limpa o filtro na coluna fornecida.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > color|A cadeia HTML de cor usada para filtrar células. Usada com a filtragem "cellColor" e "fontColor".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > criterion1|O primeiro critério usado para filtrar os dados. Usado como um operador no caso de filtragem "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > criterion2|O segundo critério usado para filtrar os dados. Só é usado como um operador no caso de filtragem "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > dynamicCriteria|Os critérios dinâmicos do conjunto Excel.DynamicFilterCriteria a serem aplicados nessa coluna. Usados com a filtragem "dynamic". Os valores possíveis são: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > filterOn|A propriedade usada pelo filtro para determinar se os valores devem ficar visíveis. Os valores possíveis são: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > operator|O operador usado para combinar o critério 1 e 2 ao usar a filtragem "custom". Os valores possíveis são: "And", "Or".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > values|O conjunto de valores a serem usados como parte da filtragem "values".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Relação_ > icon|O ícone usado para filtrar células. Usado com a filtragem "icon".|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propriedade_ > date|A data no formato ISO8601 usada para filtrar os dados.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propriedade_ > specificity|Como a data específica deve ser usada para manter os dados. Por exemplo, se a data for 2005-04-02 e a especificidade estiver definida como "mês", a operação de filtragem manterá todas as linhas com uma data do mês de abril de 2009. Os valores possíveis são: Ano, segunda-feira, dia, hora, minuto, segundo.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propriedade_ > formulaHidden|Indica se o Excel ocultará a fórmula para as células no intervalo. Um valor nulo indica que o intervalo inteiro não tem configuração uniforme de fórmula oculta.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propriedade_ > locked|Indica se o Excel bloqueia as células no objeto. Um valor nulo indica que o intervalo inteiro não tem configuração de bloqueio uniforme.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propriedade_ > index|Representa o índice do ícone no conjunto fornecido.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propriedade_ > set|Representa o conjunto do qual ícone faz parte. Os valores possíveis são: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > columnHidden|Representa se todas as colunas do intervalo atual estão ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > formulasR1C1|Representa a fórmula em notação no estilo L1C1.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > hidden|Representa se todas as células do intervalo atual estão ocultas. Somente leitura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > rowHidden|Representa se todas as linhas do intervalo atual estão ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Relação_ > sort|Representa a classificação de intervalo do intervalo atual. Somente leitura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Método_ > merge(across: bool)|Mescla as células do intervalo em uma região da planilha.|1.2|
|[range](/javascript/api/excel/excel.range)|_Método_ > unmerge()|Desfaz a mesclagem das células do intervalo em células separadas.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > columnWidth|Obtém ou define a largura de todas as colunas dentro do intervalo. Se as larguras das colunas não forem uniformes, será retornado null.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > rowHeight|Obtém ou define a altura de todas as linhas do intervalo. Se as alturas das linhas não forem uniformes, será retornado null.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Relação_ > protection|Retorna o objeto de proteção de formato para um intervalo. Somente leitura.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Método_ > autofitColumns()|Altera a largura das colunas do intervalo atual para obter o melhor ajuste, com base nos dados atuais nas colunas.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Método_ > autofitRows()|Altera a altura das linhas do intervalo atual para obter o melhor ajuste, com base nos dados atuais nas colunas.|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Propriedade_ > address|Representa as linhas visíveis do intervalo atual.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Método_apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|Executa uma operação de classificação.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > ascending|Indica se a classificação é feita de forma crescente.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > color|Representa a cor que é o destino da condição se a classificação estiver na cor da fonte ou da célula.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > dataOption|Representa as opções de classificação adicionais para esse campo. Os valores possíveis são: Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > key|Representa a coluna (ou linha, dependendo da orientação da classificação) em que a condição está. Representado como um deslocamento da primeira coluna (ou linha).|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > sortOn|Representa o tipo de classificação dessa condição. Os valores possíveis são: Valor, CellColor, FontColor, Ícone.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Relação_ > icon|Representa o ícone que é o destino da condição se a classificação está no ícone da célula.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relação_ > sort|Representa a classificação da tabela. Somente leitura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relação_ > planilha|A planilha que contém a tabela atual. Somente leitura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > clearFilters()|Limpa todos os filtros aplicados à tabela no momento.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > convertToRange()|Converte a tabela em um intervalo de células normal. Todos os dados são preservados.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > reapplyFilters()|Aplica novamente todos os filtros à tabela.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Relação_ > filter|Recupera o filtro aplicado à coluna. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propriedade_ > matchCase|Indica se o uso de maiúsculas ou minúsculas afetou a última classificação da tabela. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propriedade_ > method|Indica o último método de ordenação de caracteres chineses usado para classificar a tabela. Somente leitura. Os valores possíveis são: PinYin, StrokeCount.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Relação_ > fields|Representa as condições atuais usadas para a última classificação da tabela. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_apply(fields: SortField[], matchCase: bool, method: string)|Executa uma operação de classificação.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > clear()|Limpa a classificação que está na tabela. Essa ação não modifica a ordenação da tabela, mas limpa o estado dos botões do cabeçalho.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > reapply()|Reaplica os parâmetros de classificação atuais à tabela.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > funções|Representa uma instância de aplicativo do Excel que contém essa pasta de trabalho. Somente leitura.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > protection|Retorna o objeto de proteção da planilha para uma planilha. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Propriedade_ > protected|Indica se a planilha está protegida. Somente Leitura. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Relação_ > options|Opções de proteção da planilha. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Método_ > protect(options: WorksheetProtectionOptions)|Protege uma planilha. Falhará se uma planilha estiver protegida.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Método_ > unprotect()|Desprotege uma planilha.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowAutoFilter|Indica a opção de proteção de planilha para permitir a utilização do recurso de filtro automático.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowDeleteColumns|Indica a opção de proteção de planilha para permitir a exclusão de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowDeleteRows|Indica a opção de proteção de planilha para permitir a exclusão de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatCells|Indica a opção de proteção de planilha para permitir a formatação de células.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatColumns|Indica a opção de proteção de planilha para permitir a formatação de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatRows|Indica a opção de proteção de planilha para permitir a formatação de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertColumns|Indica a opção de proteção de planilha para permitir a inserção de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertHyperlinks|Indica a opção de proteção de planilha para permitir a inserção de hiperlinks.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertRows|Indica a opção de proteção de planilha para permitir a inserção de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowPivotTables|Indica a opção de proteção de planilha para permitir a utilização do recurso de Tabela Dinâmica.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowSort|Indica a opção de proteção de planilha para permitir a utilização do recurso de classificação.|1.2|

## <a name="excel-javascript-api-11"></a>API JavaScript do Excel 1.1

A API JavaScript do Excel 1.1 é a primeira versão da API. Para saber mais sobre a API, confira [a API JavaScript do Excel](/javascript/api/excel) nos tópicos de referência.

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
