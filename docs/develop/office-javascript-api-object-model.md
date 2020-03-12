---
title: Modelo de objeto comum de API JavaScript para Office
description: ''
ms.date: 03/10/2020
localization_priority: Normal
ms.openlocfilehash: 85ecd3b7b676a11a4ff41868adbbd9a0d907f32a
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596722"
---
# <a name="common-javascript-api-object-model"></a>Modelo de objeto comum de API JavaScript para Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Os suplementos de JavaScript do Office oferecem acesso à funcionalidade subjacente do host. A maioria desse acesso percorre alguns objetos importantes. O objeto [contexto](#context-object) oferece acesso ao tempo de execução ambiente depois de inicialização. O objeto[documento](#document-object) oferece o controle do usuário a um documento do Excel, PowerPoint ou Word. O objeto [caixa de correio](#mailbox-object) oferece um acesso ao suplemento do Outlook para mensagens e perfis de usuário. Noções básicas sobre as relações entre esses objetos gerais é a base de um suplemento JavaScript.

## <a name="context-object"></a>Objeto de contexto

**Aplica-se a:** todos os tipos de suplementos

Quando um suplemento é [inicializado](initialize-add-in.md), ele possui diversos objetos diferentes com os quais pode interagir no ambiente do tempo de execução. O contexto do tempo de execução do suplemento é refletido na API por meio do objeto [Contexto](/javascript/api/office/office.context). O**Contexto** é o principal objeto que fornece acesso aos objetos mais importantes da API, como os objetos [Documento](/javascript/api/office/office.document) e [Caixa de correio](/javascript/api/outlook/Office.mailbox) que, por sua vez, fornecem acesso ao conteúdo do documento e da caixa de correio.

Por exemplo, nos suplementos do painel de tarefas e de conteúdo, é possível usar a propriedade [documento](/javascript/api/office/office.context#document) do objeto **Context** para acessar as propriedades e os métodos do objeto **Document**. Isso permite interagir com o conteúdo de documentos do Word, planilhas do Excel ou tarefas do Project. Do mesmo modo, com os suplementos do Outlook, você pode usar a propriedade [mailbox](/javascript/api/outlook/Office.mailbox) do objeto **Context** para acessar as propriedades e os métodos do objeto **Mailbox** e interagir com a mensagem, a solicitação de reunião ou o conteúdo do compromisso.

O objeto **Contexto** também fornece acesso às propriedades [contentLanguage](/javascript/api/office/office.context#contentlanguage) e [displayLanguage](/javascript/api/office/office.context#displaylanguage) que permitem determinar a localidade (idioma) usada no documento ou no item, ou pelo aplicativo host. A propriedade [roamingSettings](/javascript/api/office/office.context#roamingsettings) permite que você acesse os membros do objeto [RoamingSettings](/javascript/api/office/office.context#roamingsettings), que armazena configurações específicas para o suplemento para caixas de correio de usuários individuais. Por fim, o objeto **Contexto** fornece uma propriedade [ui](/javascript/api/office/office.ui) que permite que o suplemento inicie caixas de diálogo pop-up.


## <a name="document-object"></a>Objeto Document

**Aplica-se a:** tipos de suplemento de conteúdo e painel de tarefas

Para interagir com dados de documento no Excel, PowerPoint e Word, a API fornece o objeto [Document](/javascript/api/office/office.document) . Você pode usar `Document` membros do objeto para acessar dados das seguintes maneiras:

- Ler e gravar as seleções ativas na forma de texto, células contíguas (matrizes) ou tabelas.

- Dados tabulares (matrizes ou tabelas).

- Associações (criadas com os métodos "Add" do `Bindings` objeto).

- Partes XML personalizadas (somente para Word).

- Configurações ou estado do suplemento persistido por suplemento no documento.

Você também pode usar o `Document` objeto para interagir com dados em documentos do projeto. A funcionalidade específica do projeto da API está documentada na classe abstrata Members [ProjectDocument](/javascript/api/office/office.document) . Para obter mais informações sobre a criação de suplementos do painel de tarefas para o Project, consulte [suplementos do painel de tarefas para o Project](../project/project-add-ins.md).

Todas essas formas de acesso a dados começam a partir de uma instância `Document` do objeto abstract.

Você pode acessar uma instância do `Document` objeto quando o painel de tarefas ou o suplemento de conteúdo é inicializado usando a propriedade [Document](/javascript/api/office/office.context#document) do `Context` objeto. O `Document` objeto define as funções de acesso de dados comuns compartilhadas entre documentos do Word e Excel, `CustomXmlParts` e também fornece acesso ao objeto para documentos do Word.

O `Document` objeto suporta quatro maneiras para os desenvolvedores acessar o conteúdo do documento:


- Acesso baseado em seleção

- Acesso baseado em associação

- Acesso baseado em partes personalizadas do XML (apenas para Word)

- Acesso baseado em documento (somente para Word e PowerPoint)

Para ajudá-lo a entender como os métodos de acesso de dados baseados na seleção e na associação funcionam, explicaremos como as APIs de acesso aos dados proporcionam acesso consistente aos dados de diferentes aplicativos do Office.


### <a name="consistent-data-access-across-office-applications"></a>Acesso consistente aos dados entre aplicativos do Office

 **Aplica-se a:** tipos de suplemento de conteúdo e painel de tarefas

Para criar extensões que funcionem perfeitamente entre diferentes documentos do Office, a API JavaScript do Office abstrai as particularidades de cada aplicativo do Office por meio de tipos de dados comuns e a capacidade de forçar o conteúdo de um documento diferente em três tipos de dados comuns.


#### <a name="common-data-types"></a>Tipo comuns de dados

Nos acessos a dados baseados em seleção e em associação, os conteúdos dos documentos são expostos por meio dos tipos de dados comuns a todos os aplicativos compatíveis do Office. No Office 2013, há suporte para três tipos de dados principais:



|**Tipo de dados**|**Descrição**|**Suporte ao aplicativo de host**|
|:-----|:-----|:-----|
|Texto|Fornece uma representação, em uma cadeia de caracteres, dos dados na seleção ou associação.|No Excel 2013, no Project 2013 e no PowerPoint 2013, há suporte apenas para texto sem formatação. No Word 2013, há suporte para três formatos de texto: texto sem formatação, HTML e OOXML (Office Open XML). Quando o texto é selecionado em uma célula no Excel, os métodos baseados em seleção realizam os processos de leitura e gravação para todo o conteúdo da célula, mesmo que apenas uma parte do texto esteja selecionada na célula. Quando texto é selecionado no Word e no PowerPoint, os métodos baseados em seleção realizam os processos de leitura e gravação apenas para os caracteres selecionados. O Project 2013 e o PowerPoint 2013 dão suporte apenas ao acesso a dados com base em seleção.|
|Matriz|Fornece os dados na seleção ou associação como uma **Array** bidimensional, que, no JavaScript, é implementada como uma matriz de matrizes. Por exemplo, duas linhas de valores **string** em duas colunas seriam ` [['a', 'b'], ['c', 'd']]`, e uma única coluna com três linhas seria `[['a'], ['b'], ['c']]`.|Há suporte ao acesso a dados de matriz apenas no Excel 2013 e no Word 2013.|
|Tabela|Fornece os dados na seleção ou associação como um objeto [TableData](/javascript/api/office/office.tabledata) . O `TableData` objeto expõe os dados por meio `headers` das `rows` Propriedades e.|Há suporte ao acesso a dados de tabela apenas no Excel 2013 e no Word 2013.|

#### <a name="data-type-coercion"></a>Coerção de tipo de dados

Os métodos de acesso a dados `Document` nos objetos de vinculação e de [Associação](/javascript/api/office/office.binding) oferecem suporte à especificação do tipo de dados desejado usando o parâmetro _coercionType_ desses métodos e valores de enumeração [coercionType](/javascript/api/office/office.coerciontype) correspondentes. Independentemente da forma real da associação, os diferentes aplicativos do Office dão suporte aos tipos de dados comuns tentando forçar os dados para o tipo de dados solicitado. Por exemplo, se uma tabela ou parágrafo do Word é selecionado, o desenvolvedor pode especificar para lê-lo como texto sem formatação, HTML, Office Open XML ou uma tabela, e a implementação da API trata das transformações e conversões de dados necessárias.


> [!TIP]
> **Quando você deve usar a matriz versus a tabela coercionType para acesso a dados?** Se você precisar que seus dados de tabulação sejam expandidos dinamicamente quando as linhas e colunas forem adicionadas, e você deve trabalhar com cabeçalhos de tabela, você deve usar o tipo de dados Table `Document` ( `Binding` especificando o parâmetro _coercionType_ de `"table"` um `Office.CoercionType.Table`objeto ou um método de acesso a dados como ou). A adição de linhas e colunas dentro da estrutura de dados é suportada em dados de tabela e matriz, mas só é possível acrescentar linhas e colunas para dados de tabela. Se você não estiver planejando a adição de linhas e colunas e seus dados não exigirem funcionalidade de cabeçalho, você deverá usar o tipo de dados Matrix (especificando o parâmetro _coercionType_ do método de acesso `"matrix"` a `Office.CoercionType.Matrix`dados como ou), que fornece um modelo mais simples de interagir com os dados.

Se os dados não puderem ser forçados para o tipo especificado, a propriedade [AsyncResult.status](/javascript/api/office/office.asyncresult#status) presente nos retornos de chamada retorna `"failed"`, e você pode usar a propriedade [AsyncResult.error](/javascript/api/office/office.asyncresult#error) para acessar um objeto [Error](/javascript/api/office/office.error) com informações sobre o motivo pelo qual a chamada de método falhou.


## <a name="working-with-selections-using-the-document-object"></a>Trabalhar com seleções que usam o objeto Document


O `Document` objeto expõe métodos que permitem que você leia e grave na seleção atual do usuário em uma maneira "Get e Set". Para fazer isso, o `Document` objeto fornece os `getSelectedDataAsync` métodos `setSelectedDataAsync` e.

Para obter exemplos de códigos que demostram como realizar tarefas com seleções, consulte [Ler e gravar dados na seleção ativa em um documento ou uma planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Trabalhar com associações usando os objetos Bindings e Binding


O acesso a dados baseado em associação habilita os suplementos de conteúdo e painel de tarefas a acessarem de forma consistente determinada região de um documento ou uma planilha por meio de um identificador vinculado a uma associação. Primeiro, o suplemento precisa estabelecer a associação chamando um dos métodos que vinculam uma parte do documento a um identificador exclusivo: [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-), [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) ou [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-). Depois que a associação é estabelecida, o suplemento pode usar o identificador fornecido para acessar os dados contidos na região vinculada do documento ou da planilha. A criação de associações fornece o seguinte valor ao suplemento:


- Permite o acesso a estruturas comuns de dados em aplicativos compatíveis do Office, como: tabelas, intervalos ou texto (uma execução contígua de caracteres).

- Habilita operações de leitura/gravação sem exigir que o usuário realize uma seleção.

- Estabelece uma relação entre o suplemento e os dados presentes no documento. As associações estão presentes no documento e podem ser acessadas em um momento posterior.

A criação de uma associação também permite que você se inscreva em eventos de alteração de seleção e de dados que apresentem um escopo definido para essa região específica do documento ou da planilha. Isso significa que o suplemento só é notificado sobre alterações que ocorrem dentro da região associada, e não sobre alterações gerais que ocorrem em todo o documento ou planilha.

O objeto [bindings](/javascript/api/office/office.bindings) expõe um método [getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) que fornece acesso ao conjunto de todas as associações estabelecidas no documento ou na planilha. Uma associação individual pode ser acessada por sua ID usando os métodos [bindings. getBindingByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) ou [Office. Select](/javascript/api/office) . Você pode estabelecer novas associações, bem como remover as existentes usando um dos seguintes `Bindings` métodos do objeto: [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-), [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-), [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-)ou [releaseByIdAsync](/javascript/api/office/office.bindings#releasebyidasync-id--options--callback-).

Há três tipos diferentes de associações que você especifica com o parâmetro _BindingType_ quando você cria uma associação com os `addFromSelectionAsync`métodos, `addFromPromptAsync` ou: `addFromNamedItemAsync`



|**Tipo de associação**|**Descrição**|**Suporte ao aplicativo de host**|
|:-----|:-----|:-----|
|Associação de texto|Associa a uma região do documento que pode ser representada como um texto.|No Word, a maioria das seleções contíguas são válidas, enquanto no Excel apenas as seleções de células únicas podem ser usadas para uma associação de texto. No Excel, só há suporte para texto sem formatação. No Word, há suporte para três formatos: texto sem formatação, HTML e Open XML do Office.|
|Associação de matriz|Associa a uma região fixa de um documento que contém dados tabulares sem cabeçalhos. Os dados de uma associação de matriz são gravados ou lidos como uma **Array** bidimensional, que é implementada como uma matriz de matrizes no JavaScript. Por exemplo, duas linhas de valores **string** em duas colunas podem ser gravadas ou lidas como ` [['a', 'b'], ['c', 'd']]`, e uma única coluna de três linhas pode ser gravada ou lida como `[['a'], ['b'], ['c']]`.|No Excel, qualquer seleção contígua de células pode ser usada para estabelecer uma associação de matriz. No Word, apenas as tabelas dão suporte à associação de matriz.|
|Associação de tabelas|Vincula a uma região de um documento que contém uma tabela com cabeçalhos. Os dados em uma vinculação de tabela são gravados ou lidos como um objeto [TableData](/javascript/api/office/office.tabledata) . O `TableData` objeto expõe os dados por meio das propriedades de **cabeçalhos** e **linhas** .|Qualquer tabela do Excel ou Word pode ser a base para uma associação de tabela. Após estabelecer uma associação de tabelas, as linhas ou colunas novas que um usuário adicionar à tabela são automaticamente incluídas na associação.  |

<br/>

Depois que `Bindings` uma associação é criada usando um dos três métodos "Add" do objeto, você pode trabalhar com as propriedades e os dados da Associação usando os métodos do objeto correspondente: [matrixbinding](/javascript/api/office/office.matrixbinding), [TableBinding](/javascript/api/office/office.tablebinding)ou [TextBinding](/javascript/api/office/office.textbinding). Todos esses três objetos herdam os métodos [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) e [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-) do `Binding` objeto que permitem que você interaja com os dados associados.

Para obter exemplos de códigos que demonstram como realizar tarefas com associações, consulte [Associar a regiões em um documento ou uma planilha](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Trabalhar com partes XML personalizadas usando os objetos CustomXmlParts e CustomXmlPart


 **Aplica-se a:** suplementos de painel de tarefas para Word

Os objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts) e [CustomXmlPart](/javascript/api/office/office.customxmlpart) da API fornecem acesso a partes XML personalizadas de documentos do Word, que permitem a manipulação orientada por XML de conteúdo do documento. Para demonstrações de como trabalhar com `CustomXmlParts` os `CustomXmlPart` objetos e, confira o exemplo de código [Word-Add-in-work-with-Custom-XML-Parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) .


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>Trabalhar com o documento inteiro usando o método getFileAsync


 **Aplica-se a:** suplementos de painel de tarefas para Word e PowerPoint

O método [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) e os membros dos objetos [File](/javascript/api/office/office.file) e [Slice](/javascript/api/office/office.slice) fornecem a funcionalidade necessária para obter documentos inteiros do Word e PowerPoint em fatias (frações) de até 4 MB por vez. Para saber mais, consulte [Obter todo o documento por meio de um suplemento para PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## <a name="mailbox-object"></a>Objeto Mailbox

**Aplica-se a:** suplementos do Outlook

Os suplementos do Outlook usam principalmente um subconjunto da API exposta no objeto [Mailbox](/javascript/api/outlook/Office.mailbox). Para acessar os objetos e membros específicos para suplementos do Outlook, como o objeto [Item](/javascript/api/outlook/Office.mailbox), use a propriedade [mailbox](/javascript/api/outlook/Office.mailbox) do objeto **Context** para acessar o objeto **Mailbox**, conforme exibido na linha de código abaixo.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Além disso, os suplementos do Outlook podem usar os seguintes objetos:

- `Office`objeto: para inicialização.

- `Context`objeto: para acessar as propriedades de conteúdo e idioma de exibição.

- `RoamingSettings`objeto: para salvar as configurações personalizadas do suplemento do Outlook na caixa de correio do usuário onde o suplemento está instalado.

Para obter informações sobre como usar o JavaScript em suplementos do Outlook, confira [Suplementos do Outlook ](../outlook/outlook-add-ins-overview.md).

