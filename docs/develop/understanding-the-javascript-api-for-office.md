---
title: No??es b?sicas da API JavaScript para Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 1ff65e8cf081330c0ce5fe8d048f703b259a5ef3
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="understanding-the-javascript-api-for-office"></a>No??es b?sicas da API JavaScript para Office

Este artigo fornece informa??es sobre a API JavaScript para Office e como us?-la. Para referenciar as informa??es, consulte [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). Para obter informa??es sobre como atualizar os arquivos de projeto do Visual Studio para a vers?o mais recente da API JavaScript para Office, consulte [Atualizar a vers?o da API JavaScript para Office e arquivos de esquema do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experi?ncia do Office depois de cri?-lo, verifique se voc? est? em conformidade com as [Pol?ticas de valida??o do AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Por exemplo, para passar na valida??o, seu suplemento deve funcionar em todas as plataformas com suporte aos m?todos que voc? definir (para mais informa??es, confira a [se??o 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [P?gina de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Fazer refer?ncia ? biblioteca da API JavaScript para Office no suplemento

A biblioteca da [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) consiste no arquivo Office.js e nos arquivos .js espec?ficos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. O m?todo mais simples de fazer refer?ncia ? API ? usando nossa CDN e adicionando o seguinte `<script>` ? marca `<head>` da sua p?gina:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Isso baixar? e colocar? os arquivos da API JavaScript para Office em cache quando o suplemento for carregado pela primeira vez a fim de garantir que o suplemento esteja usando a implementa??o mais recente do Office.js e de seus arquivos associados na vers?o especificada.

Para saber mais sobre a CDN do Office.js, incluindo como ? feito o controle de vers?o e como lidar com a compatibilidade com vers?es anteriores, consulte [Fazer refer?ncia ? biblioteca da API JavaScript para Office de sua rede de distribui??o de conte?do (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Iniciar o suplemento

**Aplica-se a:** todos os tipos de suplementos

O Office.js fornece um evento de inicializa??o que ? acionado quando a API est? totalmente carregada e pronta para come?ar a intera??o com o usu?rio. Voc? pode usar o manipulador de eventos **initialize** para implementar cen?rios comuns de inicializa??o de suplementos, como solicitar que o usu?rio selecione algumas c?lulas no Excel e, em seguida, insira um gr?fico gerado a partir desses valores selecionados. Voc? tamb?m pode usar o manipulador de eventos de inicializa??o para inicializar outras l?gicas personalizadas do suplemento, como estabelecer associa??es, solicitar valores padr?o de configura??o do suplemento e assim por diante.

No m?nimo, o evento de inicializa??o se pareceria com o exemplo a seguir:     

```js
Office.initialize = function () { };
```
Se voc? estiver usando estruturas JavaScript adicionais que incluem seus pr?prios manipuladores de inicializa??o ou testes, esses devem ser colocados dentro do evento Office.initialize. Por exemplo, a fun??o [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

Todas as p?ginas dentro de Suplementos do Office s?o necess?rias para atribuir um manipulador de eventos ao evento de inicializa??o, **Office.initialize**. Se voc? n?o incluir um manipulador de eventos, o suplemento poder? gerar um erro ao iniciar. Al?m disso, se um usu?rio tentar usar o suplemento com um cliente Web do Office Online, como o Excel Online, o PowerPoint Online ou o Outlook Web App, ele n?o funcionar?. Se voc? n?o precisar de nenhum c?digo de inicializa??o, ent?o, o corpo da fun??o atribu?da a **Office.initialize** poder? ficar vazio, como no primeiro exemplo acima.

Para obter mais detalhes sobre a sequ?ncia de eventos na inicializa??o do suplemento, consulte [Carregar o DOM e o ambiente de execu??o](loading-the-dom-and-runtime-environment.md).

#### <a name="initialization-reason"></a>Motivo da inicializa??o
Para suplementos de conte?do e de painel de tarefas, o Office.initialize fornece um par?metro _reason_ adicional. Esse par?metro pode ser usado para determinar como um suplemento foi adicionado ao documento atual. Voc? pode usar isso para fornecer l?gica diferente para quando um suplemento pela primeira vez em compara??o com quando j? existia dentro do documento. 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```
Para obter mais informa??es, confira [Evento Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) e [Enumera??o InitializationReason](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration). 

## <a name="context-object"></a>Objeto Context

**Aplica-se a:** todos os tipos de suplementos

Quando um suplemento ? iniciado, ele possui diversos objetos diferentes com os quais pode interagir no ambiente de tempo de execu??o. O contexto do tempo de execu??o do suplemento ? refletido na API por meio do objeto [Context](https://dev.office.com/reference/add-ins/shared/office.context). **Context** ? o principal objeto e fornece acesso aos objetos mais importantes da API, como [Document](https://dev.office.com/reference/add-ins/shared/document) e [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) que, por sua vez, fornecem acesso ao conte?do do documento e da caixa de correio.

Por exemplo, nos suplementos do painel de tarefas e de conte?do, ? poss?vel usar a propriedade [documento](https://dev.office.com/reference/add-ins/shared/office.context.document) do objeto **Context** para acessar as propriedades e os m?todos do objeto **Document**. Isso permite interagir com o conte?do de documentos do Word, planilhas do Excel ou tarefas do Project. Do mesmo modo, com os suplementos do Outlook, voc? pode usar a propriedade [mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) do objeto **Context** para acessar as propriedades e os m?todos do objeto **Mailbox** e interagir com a mensagem, a solicita??o de reuni?o ou o conte?do do compromisso.

O objeto **Context** tamb?m fornece acesso ?s propriedades [contentLanguage](https://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) e [displayLanguage](https://dev.office.com/reference/add-ins/shared/office.context.displaylanguage) que permitem determinar a localidade (idioma) usada no documento ou no item, ou pelo aplicativo host. E a propriedade [roamingSettings](https://dev.office.com/reference/add-ins/outlook/Office.context), que permite acessar os membros do objeto [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings). Por fim, o objeto **Context** fornece uma propriedade [ui](https://dev.office.com/reference/add-ins/shared/officeui) que permite que o suplemento inicie caixas de di?logo pop-up.


## <a name="document-object"></a>Objeto Document

**Aplica-se a:** tipos de suplemento de conte?do e painel de tarefas

Para interagir com dados do documento no Excel, PowerPoint e Word, a API fornece o objeto [Document](https://dev.office.com/reference/add-ins/shared/document). Voc? pode usar objetos membros de **Document** para acessar dados das seguintes maneiras:

- Ler e gravar as sele??es ativas na forma de texto, c?lulas cont?guas (matrizes) ou tabelas.
    
- Dados tabulares (matrizes ou tabelas).
    
- Associa??es (criadas com os m?todos "add" do objeto **Bindings**).
    
- Partes XML personalizadas (somente para Word).
    
- Configura??es ou estado do suplemento persistido por suplemento no documento.
    
Voc? tamb?m pode usar o objeto **Document** para interagir com os dados nos documentos do Project. A funcionalidade espec?fica do Project para a API est? documentada nos membros da classe abstrata [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument). Para saber mais sobre a cria??o de suplementos de painel de tarefas, consulte [Suplementos de painel de tarefas para o Project](../project/project-add-ins.md).

Todas essas formas de acesso a dados t?m in?cio em uma inst?ncia do objeto abstrato **Document**.

Voc? pode acessar uma inst?ncia do objeto **Document** quando o suplemento de painel de tarefas ou de conte?do for iniciado com a propriedade [document](https://dev.office.com/reference/add-ins/shared/office.context.document) do objeto **Context**. O objeto **Document** define fun??es comuns do acesso a dados compartilhadas em documentos do Word e do Excel, al?m de fornecer acesso ao objeto **CustomXmlParts** para documentos do Word.

O objeto **Document** permite que os desenvolvedores acessem o conte?do de documentos de quatro maneiras:


- Acesso baseado em sele??o
    
- Acesso baseado em associa??o
    
- Acesso baseado em partes personalizadas do XML (apenas para Word)
    
- Acesso baseado em documento (somente para Word e PowerPoint)
    
Para ajud?-lo a entender como os m?todos de acesso de dados baseados na sele??o e na associa??o funcionam, explicaremos como as APIs de acesso aos dados proporcionam acesso consistente aos dados de diferentes aplicativos do Office.


### <a name="consistent-data-access-across-office-applications"></a>Acesso consistente aos dados entre aplicativos do Office

 **Aplica-se a:** tipos de suplemento de conte?do e painel de tarefas

Para criar extens?es que funcionam perfeitamente em diferentes documentos do Office, a API JavaScript para Office destaca as particularidades de todos os aplicativos do Office por meio de tipos de dados comuns e da habilidade de for?ar diferentes conte?dos de documento para tr?s tipos comuns de dados.


#### <a name="common-data-types"></a>Tipo comuns de dados

Nos acessos a dados baseados em sele??o e em associa??o, os conte?dos dos documentos s?o expostos por meio dos tipos de dados comuns a todos os aplicativos compat?veis do Office. No Office 2013, h? suporte para tr?s tipos de dados principais:



|**Tipo de dados**|**Descri??o**|**Suporte ao aplicativo de host**|
|:-----|:-----|:-----|
|Texto|Fornece uma representa??o, em uma cadeia de caracteres, dos dados na sele??o ou associa??o.|No Excel 2013, no Project 2013 e no PowerPoint 2013, h? suporte apenas para texto sem formata??o. No Word 2013, h? suporte para tr?s formatos de texto: texto sem formata??o, HTML e OOXML (Office Open XML). Quando o texto ? selecionado em uma c?lula no Excel, os m?todos baseados em sele??o realizam os processos de leitura e grava??o para todo o conte?do da c?lula, mesmo que apenas uma parte do texto esteja selecionada na c?lula. Quando texto ? selecionado no Word e no PowerPoint, os m?todos baseados em sele??o realizam os processos de leitura e grava??o apenas para os caracteres selecionados. O Project 2013 e o PowerPoint 2013 d?o suporte apenas ao acesso a dados com base em sele??o.|
|Matriz|Fornece os dados na sele??o ou associa??o como uma **Array** bidimensional, que, no JavaScript, ? implementada como uma matriz de matrizes. Por exemplo, duas linhas de valores **string** em duas colunas seriam ` [['a', 'b'], ['c', 'd']]`, e uma ?nica coluna com tr?s linhas seria `[['a'], ['b'], ['c']]`.|H? suporte ao acesso a dados de matriz apenas no Excel 2013 e no Word 2013.|
|Tabela|Fornece os dados na sele??o ou associa??o como um objeto [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). O objeto **TableData** exp?e os dados por meio de propriedades **headers** e **rows**.|H? suporte ao acesso a dados de tabela apenas no Excel 2013 e no Word 2013.|

#### <a name="data-type-coercion"></a>Coer??o de tipo de dados

Os m?todos de acesso de dados nos objetos **Document** e [Binding](https://dev.office.com/reference/add-ins/shared/binding) permitem especificar o tipo desejado de dados por meio do par?metro _coercionType_ desses m?todos e os valores de enumera??o [CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) correspondentes. Independentemente da forma real da associa??o, os diferentes aplicativos do Office d?o suporte aos tipos de dados comuns ao tentar for?ar os dados a usarem o tipo de dados solicitado. Por exemplo, se uma tabela ou um par?grafo do Word for selecionado, o desenvolvedor pode escolher se deseja l?-lo como texto sem formata??o, Office Open XML ou tabela, e a implementa??o da API manipula as convers?es de dados e as transforma??es necess?rias.


> [!TIP]
> **Quando devo usar a matriz ou a tabela coercionType para o acesso aos dados?** Se for preciso que os dados tabulares cres?am dinamicamente quando linhas e colunas s?o adicionadas e voc? precisar trabalhar com os cabe?alhos da tabela, use o tipo de dados da tabela (especificando o par?metro _coercionType_ de um m?todo de acesso a dados do objeto **Document** ou **Binding** como `"table"` ou **Office.CoercionType.Table**). A adi??o de linhas e colunas na estrutura de dados tem suporte nos dados de tabela e matriz, mas o acr?scimo de linhas e colunas s? tem suporte para dados de tabela. Se voc? n?o planeja adicionar linhas e colunas, e os dados n?o exigem a funcionalidade do cabe?alho, use o tipo de dados de matriz (especificando o par?metro _coercionType_ do m?todo de acesso a dados como `"matrix"` ou **Office.CoercionType.Matrix**), que fornece um modelo mais simples para interagir com os dados.

Se os dados n?o puderem ser for?ados para o tipo especificado, a propriedade [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) presente nos retornos de chamada retorna `"failed"`, e voc? pode usar a propriedade [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) para acessar um objeto [Error](https://dev.office.com/reference/add-ins/shared/error) com informa??es sobre o motivo pelo qual a chamada de m?todo falhou.


## <a name="working-with-selections-using-the-document-object"></a>Trabalhar com sele??es que usam o objeto Document


O objeto **Document** exp?e m?todos que permitem ler e gravar a sele??o atual do usu?rio de uma maneira "obter e definir". Para fazer isso, o objeto **Document** fornece os m?todos **getSelectedDataAsync** e **setSelectedDataAsync**.

Para obter exemplos de c?digos que demostram como realizar tarefas com sele??es, consulte [Ler e gravar dados na sele??o ativa em um documento ou uma planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Trabalhar com associa??es usando os objetos Bindings e Binding


O acesso a dados baseado em associa??o habilita os suplementos de conte?do e painel de tarefas a acessarem de forma consistente determinada regi?o de um documento ou uma planilha por meio de um identificador vinculado a uma associa??o. Primeiro, o suplemento precisa estabelecer a associa??o chamando um dos m?todos que vinculam uma parte do documento a um identificador exclusivo: [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) ou [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync). Depois que a associa??o ? estabelecida, o suplemento pode usar o identificador fornecido para acessar os dados contidos na regi?o vinculada do documento ou da planilha. A cria??o de associa??es fornece o seguinte valor ao suplemento:


- Permite o acesso a estruturas comuns de dados em aplicativos compat?veis do Office, como: tabelas, intervalos ou texto (uma execu??o cont?gua de caracteres).
    
- Habilita opera??es de leitura/grava??o sem exigir que o usu?rio realize uma sele??o.
    
- Estabelece uma rela??o entre o suplemento e os dados presentes no documento. As associa??es est?o presentes no documento e podem ser acessadas em um momento posterior.
    
A cria??o de uma associa??o tamb?m permite que voc? se inscreva em eventos de altera??o de sele??o e de dados que apresentem um escopo definido para essa regi?o espec?fica do documento ou da planilha. Isso significa que o suplemento s? ? notificado sobre altera??es que ocorrem dentro da regi?o associada, e n?o sobre altera??es gerais que ocorrem em todo o documento ou planilha.

O objeto [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) exp?e um m?todo [getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync), que d? acesso ao conjunto de todas as associa??es estabelecidas no documento ou planilha. Uma associa??o individual pode ser acessada por sua ID. Para isso, use os m?todos [Bindings.getBindingByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) ou [Office.select](https://dev.office.com/reference/add-ins/shared/office.select). Voc? pode estabelecer novas associa??es e remover as associa??es existentes usando um dos seguintes m?todos para o objeto **Bindings**: [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync), [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync) ou [releaseByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync).

H? tr?s tipos diferentes de associa??es que podem ser especificadas com o par?metro _bindingType_ durante a cria??o de uma nova associa??o com os m?todos **addFromSelectionAsync**, **addFromPromptAsync** ou **addFromNamedItemAsync**:



|**Tipo de associa??o**|**Descri??o**|**Suporte ao aplicativo de host**|
|:-----|:-----|:-----|
|Associa??o de texto|Associa a uma regi?o do documento que pode ser representada como um texto.|No Word, a maioria das sele??es cont?guas s?o v?lidas, enquanto no Excel apenas as sele??es de c?lulas ?nicas podem ser usadas para uma associa??o de texto. No Excel, s? h? suporte para texto sem formata??o. No Word, h? suporte para tr?s formatos: texto sem formata??o, HTML e Open XML do Office.|
|Associa??o de matriz|Associa a uma regi?o fixa de um documento que cont?m dados tabulares sem cabe?alhos. Os dados de uma associa??o de matriz s?o gravados ou lidos como uma **Array** bidimensional, que ? implementada como uma matriz de matrizes no JavaScript. Por exemplo, duas linhas de valores **string** em duas colunas podem ser gravadas ou lidas como ` [['a', 'b'], ['c', 'd']]`, e uma ?nica coluna de tr?s linhas pode ser gravada ou lida como `[['a'], ['b'], ['c']]`.|No Excel, qualquer sele??o cont?gua de c?lulas pode ser usada para estabelecer uma associa??o de matriz. No Word, apenas as tabelas d?o suporte ? associa??o de matriz.|
|Associa??o de tabelas|Associa a uma regi?o de um documento que cont?m uma tabela com cabe?alhos. Os dados em uma associa??o de tabela s?o gravados ou lidos como um objeto [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). O objeto **TableData** exp?e os dados por meio das propriedades **headers** e **rows**.|Qualquer tabela do Excel ou Word pode ser a base para uma associa??o de tabela. Ap?s estabelecer uma associa??o de tabelas, as linhas ou colunas novas que um usu?rio adicionar ? tabela s?o automaticamente inclu?das na associa??o. |

<br/>

Depois que uma associa??o ? criada usando um dos tr?s m?todos "add" do objeto **Bindings**, ? poss?vel trabalhar com os dados e as propriedades da associa??o usando os m?todos do objeto correspondente: [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding), [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding) ou [TextBinding](https://dev.office.com/reference/add-ins/shared/binding.textbinding). Esses tr?s objetos herdam os m?todos [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) e [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync) do objeto **Binding**, o que o habilita a interagir com os dados associados.

Para obter exemplos de c?digos que demonstram como realizar tarefas com associa??es, consulte [Associar a regi?es em um documento ou uma planilha](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Trabalhar com partes XML personalizadas usando os objetos CustomXmlParts e CustomXmlPart


 **Aplica-se a:** suplementos de painel de tarefas para Word

Os objetos [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) e [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) da API fornecem acesso a partes XML personalizadas em documentos do Word, que habilitam a manipula??o orientada por XML do conte?do do documento. Para obter demonstra??es de como trabalhar com objetos **CustomXmlParts** e **CustomXmlPart**, consulte o exemplo de c?digo [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts).


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>Trabalhar com o documento inteiro usando o m?todo getFileAsync


 **Aplica-se a:** suplementos de painel de tarefas para Word e PowerPoint

O m?todo [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync) e os membros dos objetos [File](https://dev.office.com/reference/add-ins/shared/file) e [Slice](https://dev.office.com/reference/add-ins/shared/slice) fornecem a funcionalidade necess?ria para obter documentos inteiros do Word e PowerPoint em fatias (fra??es) de at? 4 MB por vez. Para saber mais, consulte [Obter todo o documento por meio de um suplemento para PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## <a name="mailbox-object"></a>Objeto Mailbox


 **Aplica-se a:** suplementos do Outlook

Os suplementos do Outlook usam principalmente um subconjunto da API exposta no objeto [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox). Para acessar os objetos e membros espec?ficos para suplementos do Outlook, como o objeto [Item](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item), use a propriedade [mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) do objeto **Context** para acessar o objeto **Mailbox**, conforme exibido na linha de c?digo abaixo.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Al?m disso, os suplementos do Outlook podem usar os seguintes objetos:


-  Objeto **Office**: para inicializa??o.
    
-  Objeto **Context**: para acesso a propriedades de conte?do e idioma de exibi??o.
    
-  Objeto **RoamingSettings**: para salvar as configura??es personalizadas do suplemento do Outlook na caixa de correio do usu?rio em que o suplemento est? instalado.
    
Para obter informa??es sobre como usar o JavaScript em suplementos do Outlook, confira [Suplementos do Outlook ](https://docs.microsoft.com/en-us/outlook/add-ins/).


## <a name="api-support-matrix"></a>Matriz de suporte da API


Esta tabela resume a API e os recursos compat?veis com os tipos de suplemento (conte?do, painel de tarefas e Outlook) e os aplicativos do Office que podem hosped?-los quando o usu?rio especifica os aplicativos hospedados pelo Office compat?veis com o suplemento usando o [esquema 1.1 do manifesto de suplementos e recursos compat?veis com a v1.1 da API JavaScript para Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Nome do host**|Banco de dados|Pasta de trabalho|Caixa de correio|Apresenta??o|Documento|Project|
||**Aplicativos host** **compat?veis**|Aplicativos Web do Access|Excel,<br/>Excel Online|Outlook,<br/>Outlook Web App,<br/>OWA para dispositivos|PowerPoint,<br/>PowerPoint Online|Word|Project|
|**Tipos de suplemento com suporte**|Conte?do|S|S||S|||
||Painel de tarefas||S||S|S|S|
||Outlook|||S||||
|**Recursos da API compat?veis**|Ler/gravar texto||S||S|S|S<br/>(Somente leitura)|
||Ler/gravar matriz||S|||S||
||Ler/gravar tabela||S|||S||
||Ler/gravar HTML|||||S||
||Leitura/grava??o<br/>Office Open XML|||||S||
||Ler propriedades de tarefa, recurso, modo de exibi??o e campo||||||S|
||Eventos alterados pela sele??o||S|||S||
||Obter documento inteiro||||S|S||
||Associa??es e eventos de associa??o|S<br/>(Somente vincula??es de tabela totais e parciais)|S|||S||
||Ler/gravar partes XML personalizadas|||||S||
||Persistir dados de estado de suplemento (configura??es)|S<br/>(Por suplemento do host)|S<br/>(Por documento)|S<br/>(Por caixa de correio)|S<br/>(Por documento)|S<br/>(Por documento)||
||Eventos alterados pelas configura??es|S|S||S|S||
||Obter o modo de exibi??o ativo<br/>e visualizar eventos alterados||||S|||
||Navegar para locais<br/>no documento||S||S|S||
||Ativar contextualmente<br/>usando regras e RegEx|||S||||
||Ler propriedades do item|||S||||
||Ler perfil de usu?rio|||S||||
||Obter anexos|||S||||
||Obter o token de identidade do usu?rio|||S||||
||Chamar os servi?os Web do Exchange|||S||||
