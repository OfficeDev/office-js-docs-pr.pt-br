---
title: Conjuntos de requisitos da API Comum do Office
description: Saiba mais sobre os conjuntos de requisitos Office API comum.
ms.date: 04/28/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 959f03bf41496c1506087c2851efad336cdec676
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253344"
---
# <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

> [!TIP]
> Procurando os conjuntos *de requisitos* de API específicos do aplicativo? Confira os seguintes conjuntos de requisitos de API:
>
> - [Conjuntos de requisitos de API JavaScript para Excel](excel-api-requirement-sets.md) (ExcelApi)
> - [Conjuntos de requisitos de API JavaScript para Word](word-api-requirement-sets.md) (WordApi)
> - [Conjuntos de requisitos de API JavaScript para OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
> - [Conjuntos de requisitos da API JavaScript do PowerPoint](powerpoint-api-requirement-sets.md) (PowerPointApi)
> - [Noções básicas sobre os conjuntos de requisitos da API do Outlook](outlook-api-requirement-sets.md) (Caixa de Correio)

> [!IMPORTANT]
> Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

## <a name="common-api-requirement-sets"></a>Conjuntos de requisitos da API Comum

As seções a seguir listam os conjuntos de requisitos da API comum, os métodos em cada conjunto e os Office cliente que suportam esse conjunto de requisitos. Todos esses conjuntos de requisitos da API são versão 1.1, a menos que especificado de outra forma.

> [!TIP]
> Precisa de informações sobre onde os complementos e conjuntos de requisitos têm suporte Office aplicativo e versão? Consulte [Office disponibilidade de aplicativo cliente e plataforma para Office de complementos](../../overview/office-add-in-availability.md).

### <a name="activeview"></a>ActiveView

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| PowerPoint no Windows<br>PowerPoint Online<br>PowerPoint no iPad<br>PowerPoint no Mac|Document.getActiveViewAsync|

---

### <a name="addincommands"></a>AddInCommands

Confira [Conjuntos de requisitos de comandos de suplementos](add-in-commands-requirement-sets.md).

---

### <a name="bindingevents"></a>BindingEvents

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access<br>Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Binding.addHandlerAsync<br>Binding.removeHandlerAsync|

---

### <a name="compressedfile"></a>CompressedFile

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel 2016 e posteriormente Windows<br>Excel Online<br>Excel 2016 e posterior no Mac<br>PowerPoint no Windows<br>PowerPoint Online<br>PowerPoint no iPad<br>PowerPoint no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Dá suporte à saída para o formato OOXML (Office Open XML) como uma matriz de bytes<br>(Office.FileType.Compressed) ao usar o método Document.getFileAsync.|

---

### <a name="customxmlparts"></a>CustomXmlParts

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|

---

### <a name="dialogapi"></a>DialogApi

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Confira [Conjuntos de requisitos da API da Caixa de Diálogo](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |

---

### <a name="documentevents"></a>DocumentEvents

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>OneNote Online<br>PowerPoint no Windows<br>PowerPoint Online<br>PowerPoint no iPad<br>PowerPoint no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|

---

### <a name="file"></a>Arquivo

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>PowerPoint no Windows<br>PowerPoint Online<br>PowerPoint no iPad<br>PowerPoint no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|

---

### <a name="htmlcoercion"></a>HtmlCoercion

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| OneNote Online<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Dá suporte à coerção para HTML (Office.CoercionType.Html) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="identityapi"></a>IdentityAPI

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Confira [Conjuntos de requisitos da API de Identidade](identity-api-requirement-sets.md). | Auth.getAccessToken |

---

### <a name="imagecoercion"></a>ImageCoercion

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Confira [conjuntos de requisitos de Coerção de Imagens](image-coercion-requirement-sets.md). | Método Document.setSelectedDataAsync|

---

### <a name="mailbox"></a>Caixa de correio

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
|Outlook no Windows<br>Outlook Online<br>Outlook no Android<br>Outlook no Mac<br>Outlook no iOS|Confira [Noções básicas sobre conjuntos de requisitos da API do Outlook](outlook-api-requirement-sets.md).|

---

### <a name="matrixbindings"></a>MatrixBindings

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>Word no Windows<br>Word Online<br>Word no iPad<br>Word no Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="matrixcoercion"></a>MatrixCoercion

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Dá suporte à coerção para a estrutura de dados “matrix” (matriz de matrizes) (Office.CoercionType.Matrix) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="ooxmlcoercion"></a>OoxmlCoercion

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Dá suporte à coerção para o formato OOXML (Open Office XML) (Office.CoercionType.Ooxml) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="openbrowserwindowapi"></a>OpenBrowserWindowApi

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Consulte [Open Browser Window API requirement sets](open-browser-window-api-requirement-sets.md). | Office.context.ui.openBrowserWindow |

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access||

---

### <a name="pdffile"></a>PdfFile

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel no Windows<br>Excel Online<br>Excel no Mac<br>PowerPoint no Windows<br>PowerPoint Online<br>PowerPoint no iPad<br>PowerPoint no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online|Dá suporte à saída para o formato PDF (Office.FileType.Pdf)<br>ao usar o método Document.getFileAsync.|

---

### <a name="ribbonapi"></a>RibbonApi

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Consulte [Conjuntos de requisitos da API da Faixa de Opções](ribbon-api-requirement-sets.md). | Office.ribbon.requestUpdate |

---

### <a name="selection"></a>Selection

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>PowerPoint no Windows<br>PowerPoint Online<br>PowerPoint no iPad<br>PowerPoint no Mac<br>Project no Windows<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Configurações

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access<br>Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>OneNote Online<br>PowerPoint no Windows<br>PowerPoint Online<br>PowerPoint no iPad<br>PowerPoint no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="sharedruntime"></a>SharedRuntime

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Consulte [Conjuntos de requisitos de tempo de execução compartilhados.](shared-runtime-requirement-sets.md) | Office.addin.getStartupBehavior<br>Office.addin.hide<br>Office.addin.onVisibilityModeChanged<br>Office.addin.setStartupBehavior<br>Office.addin.showAsTaskpane<br> |

---

### <a name="tablebindings"></a>TableBindings

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access<br>Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.addColumnsAsync<br>Binding.addRowsAsync<br>Binding.deleteAllDataValuesAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="tablecoercion"></a>TableCoercion

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access<br>Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Dá suporte à coerção para a estrutura de dados “table” (Office.CoercionType.Table) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="textbindings"></a>TextBindings

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel no Windows<br>Excel Online<br>Excel no iPad<br>Excel no Mac<br>Word 2013 e posterior e Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="textcoercion"></a>TextCoercion

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel no Windows<br>Excel Online<br>Excel no iPad<br>OneNote Online<br>PowerPoint no Windows<br>PowerPoint Online<br>PowerPoint no iPad<br>PowerPoint no Mac<br>Project no Windows<br>Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Dá suporte à coerção para o formato de texto (Office.CoercionType.Text) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="textfile"></a>TextFile

|**Aplicativos do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Word 2013 e posterior no Windows<br>Word 2016 e posterior no Mac<br>Word Online<br>Word no iPad|Dá suporte à saída para o formato de texto (Office.FileType.Text) ao usar o método Document.getFileAsync.|

---

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Métodos que não fazem parte de um conjunto de requisitos

Os métodos a seguir na OFFICE JavaScript não fazem parte de um conjunto de requisitos. Se o suplemento exigir qualquer um desses métodos, use os elementos **Methods** e **Method** no manifesto do suplemento para declarar que eles são exigidos, ou então execute a verificação de tempo de execução usando uma instrução`if`. Para obter mais informações, consulte [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).

|**Nome do método**|**Office suporte a aplicativos**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Acesse aplicativos web, Excel no Windows, Excel Online, Excel no iPad e Excel no Mac|
|Document.getFilePropertiesAsync|Excel no Windows, Excel Online, Excel no iPad, Excel no Mac, PowerPoint no Windows, PowerPoint Online, PowerPoint no iPad, PowerPoint no Mac, Word no Windows, Word Online, Word no iPad e Word no Mac|
|Document.getProjectFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013 e Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013 e Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.goToByIdAsync|Excel no Windows, Excel Online, Excel no iPad, Excel no Mac, PowerPoint no Windows, PowerPoint Online, PowerPoint no iPad, PowerPoint no Mac, Word no Windows, Word Online, Word no iPad e Word no Mac|
|Settings.addHandlerAsync|Acesse aplicativos web e Excel Online|
|Settings.refreshAsync|Acesse aplicativos web, Excel no Windows, Excel Online, PowerPoint no Windows, PowerPoint Online, Word e Word Online|
|Settings.removeHandlerAsync|Acesse aplicativos web e Excel Online|
|TableBinding.clearFormatsAsync|Excel no Windows, Excel Online, Excel no iPad e Excel no Mac|
|TableBinding.setFormatsAsync|Excel no Windows, Excel Online, Excel no iPad e Excel no Mac|
|TableBinding.setTableOptionsAsync|Excel no Windows, Excel Online, Excel no iPad e Excel no Mac|

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
