# <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comuns da API do Office

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação em tempo de execução para determinar se um host do Office oferece suporte a APIs que um suplemento precisa. Para obter mais informações, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Precisa de informações sobre onde os suplementos são suportados pelo host do Office? Confira  [Disponibilidade de host e plataforma para suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).

Procurando pelos conjuntos de requisitos de API *específicos do host*? Veja os seguintes conjuntos de API:
 
- [Conjuntos de requisitos de API JavaScript do Excel](excel-api-requirement-sets.md) (ExcelApi)
- [Conjuntos de requisitos de API JavaScript para Word](word-api-requirement-sets.md) (WordApi)
- [Conjuntos de requisitos da API JavaScript do OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
- [Noções básicas sobre conjuntos de requisitos da API do Outlook](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> Não recomendamos mais que você crie e use aplicativos da Web e bancos de dados do Access no SharePoint. Como alternativa, é recomendável que você use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para web e dispositivos móveis.

## <a name="common-api-requirement-sets"></a>Conjuntos de requisitos da API comum

A tabela a seguir lista os conjuntos de requisitos comuns da API, os métodos em cada conjunto e os aplicativos host do Office que aceitam esse conjunto de requisitos. Todos esses conjuntos de requisitos da API são versão 1.1.

|**Conjunto de requisitos**|**Host do Office**|**Métodos no conjunto**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac|Document.getActiveViewAsync|
| AddInCommands | Confira [Conjuntos de requisitos de suplementos de comandos](add-in-commands-requirement-sets.md). | |
| BindingEvents  | Aplicativos Web do Access<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Suporta saída para o formato Office Open XML (OOXML) como uma matriz de bytes<br>(Office.FileType.Compressed) ao usar o método Document.getFileAsync.|
| CustomXmlParts    | Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogApi | Confira [Conjuntos de requisitos da API de caixa de diálogo](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Arquivo  | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | OneNote Online<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para HTML (Office.CoercionType.Html) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync<br>Métodos Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| IdentityAPI | Confira [Conjuntos de requisitos de APIs de identidade](identity-api-requirement-sets.md). | Auth.getAccessTokenAsync |
| ImageCoercion | Excel<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à conversão para uma imagem (Office.CoercionType.Image) ao gravar dados usando o método Document.setSelectedDataAsync.|
| Mailbox   |Outlook para Windows<br>Outlook para Web<br>Outlook para Android<br>Outlook para Mac<br>Aplicativo Web do Outlook |Confira [Noções básicas sobre conjuntos de requisitos da API do Outlook](outlook-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word<br>Word Online<br>Word para iPad<br>Word para Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para a estrutura de dados “matrix” (matriz de matrizes) (Office.CoercionType.Matrix) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| OoxmlCoercion | Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para o formato OOXML (Open Office XML) (Office.CoercionType.Ooxml) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| PartialTableBindings  | Aplicativos Web do Access||
| PdfFile   | Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à saída para o formato PDF (Office.FileType.Pdf)<br>ao usar o método Document.getFileAsync.|
| Seleção | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Project<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Configurações  | Aplicativos Web do Access<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Aplicativos Web do Access<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Aplicativos Web do Access<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para a estrutura de dados “table” (Office.CoercionType.Table) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| TextBindings  | Excel<br>Excel Online<br>Excel para iPad<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>Excel para iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Project<br>Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à coerção para o formato de texto (Office.CoercionType.Text) ao ler e gravar dados usando os métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| TextFile  | Word 2013 e posterior<br>Word 2016 para Mac e posterior<br>Word Online<br>Word para iPad|Dá suporte à saída para o formato de texto (Office.FileType.Text) ao usar o método Document.getFileAsync.|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Métodos que não fazem parte de um conjunto de requisitos

Os seguintes métodos na API JavaScript para Office não são parte de um conjunto de requisitos. Se o suplemento exigir qualquer um desses métodos, use os elementos de **Methods** e **Method** no manifesto do suplemento para declarar que são necessários ou executar a verificação de tempo de execução usando uma instrução `if` . Para saber mais , confira [Especificar hosts do Office e requisitos de API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

|**Nome do método**|**Suporte ao host do Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Acessar aplicativos da web, Excel, Excel Online e Excel para iPad|
|Document.getFilePropertiesAsync|Excel, Excel Online, Excel para iPad, Excel para Mac, PowerPoint, PowerPoint Online, PowerPoint para iPad, PowerPoint para Mac, Word, Word Online, Word para iPad e Word para Mac|
|Document.getProjectFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 e Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013 e Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013 e Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 e Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Excel para iPad, Excel para Mac, PowerPoint, PowerPoint Online, PowerPoint para iPad, PowerPoint para Mac, Word, Word Online, Word para iPad e Word para Mac|
|Settings.addHandlerAsync|Acessar aplicativos da Web, Excel, Excel Online, PowerPoint, PowerPoint Online, Word e Word Online|
|Settings.refreshAsync|Acessar aplicativos da Web, Excel, Excel Online, PowerPoint, PowerPoint Online, Word e Word Online|
|Settings.removeHandlerAsync|Acessar aplicativos da Web, Excel, Excel Online, PowerPoint, PowerPoint Online, Word e Word Online|
|TableBinding.clearFormatsAsync|Excel, Excel Online e Excel para Mac|
|TableBinding.setFormatsAsync|Excel, Excel Online, Excel para iPad e Excel para Mac|
|TableBinding.setTableOptionsAsync|Excel, Excel Online, Excel para iPad e Excel para Mac|

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos de API e hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
