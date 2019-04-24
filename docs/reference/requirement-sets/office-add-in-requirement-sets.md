---
title: Conjuntos de requisitos da API Comum do Office
description: ''
ms.date: 04/10/2019
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 16f77e81d149aa2323760013f64fbf36f4ce7d8f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450104"
---
# <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API Comum do Office

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Precisa de informações sobre onde os suplementos têm suporte do host do Office? Consulte [Disponibilidade de host e plataforma para suplementos do Office](/office/dev/add-ins/overview/office-add-in-availability).

Procurando pelos conjuntos de requisitos de API *específicos do host*? Confira os seguintes conjuntos de requisitos de API:

- [Conjuntos de requisitos de API JavaScript para Excel](excel-api-requirement-sets.md) (ExcelApi)
- [Conjuntos de requisitos de API JavaScript para Word](word-api-requirement-sets.md) (WordApi)
- [Conjuntos de requisitos de API JavaScript para OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
- [Noções básicas sobre os conjuntos de requisitos da API do Outlook](outlook-api-requirement-sets.md) (Caixa de Correio)

> [!IMPORTANT]
> Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

## <a name="common-api-requirement-sets"></a>Conjuntos de requisitos da API Comum

As seções a seguir listam os conjuntos de requisitos da API Comum, os métodos em cada conjunto e os aplicativos host do Office compatíveis com esse conjunto de requisitos. Todos esses conjuntos de requisitos da API são versão 1.1.

### <a name="activeview"></a>ActiveView

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac|Document.getActiveViewAsync|

---

### <a name="addincommands"></a>AddInCommands

Confira [Conjuntos de requisitos de comandos de suplementos](add-in-commands-requirement-sets.md).

---

### <a name="bindingevents"></a>BindingEvents

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|

---

### <a name="compressedfile"></a>CompressedFile

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à saída para o formato OOXML (Office Open XML) como uma matriz de bytes<br>(Office.FileType.Compressed) ao usar o método Document.getFileAsync.|

---

### <a name="customxmlparts"></a>CustomXmlParts

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|

---

### <a name="dialogapi"></a>DialogApi

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Confira [Conjuntos de requisitos da API da Caixa de Diálogo](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |

---

### <a name="documentevents"></a>DocumentEvents

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|

---

### <a name="file"></a>Arquivo

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|

---

### <a name="htmlcoercion"></a>HtmlCoercion

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| OneNote Online<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para HTML (Office.CoercionType.Html) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync<br>Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="identityapi"></a>IdentityAPI

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Confira [Conjuntos de requisitos da API de Identidade](identity-api-requirement-sets.md). | Auth.getAccessTokenAsync |

---

### <a name="imagecoercion"></a>ImageCoercion

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à conversão para uma imagem (Office.CoercionType.Image) ao gravar dados usando o método Document.setSelectedDataAsync.|

---

### <a name="mailbox"></a>Caixa de correio

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
|Outlook para Windows<br>Outlook para Web<br>Outlook para Android<br>Outlook para Mac<br>Aplicativo Web do Outlook |Confira [Noções básicas sobre conjuntos de requisitos da API do Outlook](outlook-api-requirement-sets.md).|

---

### <a name="matrixbindings"></a>MatrixBindings

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word<br>Word Online<br>Word para iPad<br>Word para Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="matrixcoercion"></a>MatrixCoercion

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para a estrutura de dados “matrix” (matriz de matrizes) (Office.CoercionType.Matrix) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="ooxmlcoercion"></a>OoxmlCoercion

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para o formato OOXML (Open Office XML) (Office.CoercionType.Ooxml) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access||

---

### <a name="pdffile"></a>PdfFile

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à saída para o formato PDF (Office.FileType.Pdf)<br>ao usar o método Document.getFileAsync.|

---

### <a name="selection"></a>Seleção

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Projeto<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Configurações

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="tablebindings"></a>TableBindings

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.addColumnsAsync<br>Binding.addRowsAsync<br>Binding.deleteAllDataValuesAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="tablecoercion"></a>TableCoercion

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Aplicativos Web do Access<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para a estrutura de dados “table” (Office.CoercionType.Table) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="textbindings"></a>TextBindings

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="textcoercion"></a>TextCoercion

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Excel<br>Excel Online<br>Excel para iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Projeto<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para o formato de texto (Office.CoercionType.Text) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="textfile"></a>TextFile

|**Hosts do Office**|**Métodos no conjunto**|
|:-----|:-----|
| Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à saída para o formato de texto (Office.FileType.Text) ao usar o método Document.getFileAsync.|

---

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Métodos que não fazem parte de um conjunto de requisitos

Os seguintes métodos da API JavaScript para Office não fazem parte de um conjunto de requisitos. Se o suplemento exigir qualquer um desses métodos, use os elementos **Methods** e **Method** no manifesto do suplemento para declarar que eles são exigidos, ou então execute a verificação de tempo de execução usando uma instrução`if`. Para saber mais, confira [Especificar requisitos de API e hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

|**Nome do método**|**Suporte ao host do Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Aplicativos web do Access, Excel, Excel Online, Excel para iPad e Excel para Mac|
|Document.getFilePropertiesAsync|Excel, Excel Online, Excel para iPad, Excel para Mac, PowerPoint, PowerPoint Online, PowerPoint para iPad, PowerPoint para Mac, Word, Word Online, Word para iPad e Word para Mac|
|Document.getProjectFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013 e Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013 e Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Excel para iPad, Excel para Mac, PowerPoint, PowerPoint Online, PowerPoint para iPad, PowerPoint para Mac, Word, Word Online, Word para iPad e Word para Mac|
|Settings.addHandlerAsync|Aplicativos Web do Access e Excel Online|
|Settings.refreshAsync|Aplicativos Web do Access, Excel, Excel Online, PowerPoint, PowerPoint Online, Word e Word Online|
|Settings.removeHandlerAsync|Aplicativos Web do Access e Excel Online|
|TableBinding.clearFormatsAsync|Excel, Excel Online, Excel para iPad e Excel para Mac|
|TableBinding.setFormatsAsync|Excel, Excel Online, Excel para iPad e Excel para Mac|
|TableBinding.setTableOptionsAsync|Excel, Excel Online, Excel para iPad e Excel para Mac|

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
