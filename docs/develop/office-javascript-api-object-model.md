---
title: Modelo de objeto comum de API JavaScript para Office
description: Saiba mais sobre o Office de objeto da API comum JavaScript
ms.date: 07/08/2021
ms.localizationpriority: medium
---

# <a name="common-javascript-api-object-model"></a>Modelo de objeto comum de API JavaScript para Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office APIs JavaScript dão acesso Office funcionalidade subjacente do aplicativo cliente. A maioria desse acesso percorre alguns objetos importantes. O objeto [contexto](#context-object) oferece acesso ao tempo de execução ambiente depois de inicialização. O objeto[documento](#document-object) oferece o controle do usuário a um documento do Excel, PowerPoint ou Word. O [objeto Mailbox](#mailbox-object) fornece um Outlook de usuário para mensagens, compromissos e perfis de usuário. Compreender as relações entre esses objetos de alto nível é a base de um Office Add-in.

## <a name="context-object"></a>Objeto de contexto

**Aplica-se a:** todos os tipos de suplementos

Quando um suplemento é [inicializado](initialize-add-in.md), ele possui diversos objetos diferentes com os quais pode interagir no ambiente do tempo de execução. O contexto do tempo de execução do suplemento é refletido na API por meio do objeto [Contexto](/javascript/api/office/office.context). O **Contexto** é o principal objeto que fornece acesso aos objetos mais importantes da API, como os objetos [Documento](/javascript/api/office/office.document) e [Caixa de correio](/javascript/api/outlook/office.mailbox) que, por sua vez, fornecem acesso ao conteúdo do documento e da caixa de correio.

Por exemplo, nos suplementos do painel de tarefas e de conteúdo, é possível usar a propriedade [documento](/javascript/api/office/office.context#office-office-context-document-member) do objeto **Context** para acessar as propriedades e os métodos do objeto **Document**. Isso permite interagir com o conteúdo de documentos do Word, planilhas do Excel ou tarefas do Project. Do mesmo modo, com os suplementos do Outlook, você pode usar a propriedade [mailbox](/javascript/api/office/office.context#office-office-context-mailbox-member) do objeto **Context** para acessar as propriedades e os métodos do objeto **Mailbox** e interagir com a mensagem, a solicitação de reunião ou o conteúdo do compromisso.

O objeto **Context** também fornece acesso às propriedades [contentLanguage](/javascript/api/office/office.context#office-office-context-contentlanguage-member) e [displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) que permitem determinar a localidade (idioma) usada no documento ou item ou pelo aplicativo Office. A propriedade [roamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member) permite que você acesse os membros do objeto [RoamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member), que armazena configurações específicas para o suplemento para caixas de correio de usuários individuais. Por fim, o objeto **Contexto** fornece uma propriedade [ui](/javascript/api/office/office.context#office-office-context-ui-member) que permite que o suplemento inicie caixas de diálogo pop-up.

## <a name="document-object"></a>Objeto Document

**Aplica-se a:** tipos de suplemento de conteúdo e painel de tarefas

Para interagir com dados do documento no Excel, PowerPoint e Word, a API fornece o objeto [Document](/javascript/api/office/office.document). Você pode usar membros `Document` do objeto para acessar dados das seguintes maneiras.

- Ler e gravar as seleções ativas na forma de texto, células contíguas (matrizes) ou tabelas.

- Dados tabulares (matrizes ou tabelas).

- Vinculações (criadas com os métodos "add" do `Bindings` objeto).

- Partes XML personalizadas (somente para Word).

- Configurações ou estado do suplemento persistido por suplemento no documento.

Você também pode usar o objeto `Document` para interagir com dados em Project documentos. A funcionalidade específica do Project para a API está documentada nos membros da classe abstrata [ProjectDocument](/javascript/api/office/office.document). Para saber mais sobre a criação de suplementos de painel de tarefas, consulte [Suplementos de painel de tarefas para o Project](../project/project-add-ins.md).

Todas essas formas de acesso a dados começam de uma instância do objeto `Document` abstrato.

Você pode acessar uma instância do `Document` objeto quando o painel de tarefas ou o complemento de conteúdo é inicializado usando a propriedade [document](/javascript/api/office/office.context#office-office-context-document-member) do `Context` objeto. O `Document` objeto define funções comuns de acesso a dados compartilhadas entre documentos do Word e Excel e `CustomXmlParts` também fornece acesso ao objeto para documentos do Word.

O `Document` objeto oferece suporte a quatro maneiras para os desenvolvedores acessarem o conteúdo do documento.

- Acesso baseado em seleção

- Acesso baseado em associação

- Acesso baseado em partes personalizadas do XML (apenas para Word)

- Acesso baseado em documento (somente para Word e PowerPoint)

Para ajudá-lo a entender como os métodos de acesso de dados baseados na seleção e na associação funcionam, explicaremos como as APIs de acesso aos dados proporcionam acesso consistente aos dados de diferentes aplicativos do Office.

### <a name="consistent-data-access-across-office-applications"></a>Acesso consistente aos dados entre aplicativos do Office

 **Aplica-se a:** tipos de suplemento de conteúdo e painel de tarefas

Para criar extensões que funcionam perfeitamente em diferentes documentos Office, Office API JavaScript do Office abstrai as particularidades de cada aplicativo Office por meio de tipos de dados comuns e a capacidade de coagir conteúdos de documentos diferentes em três tipos de dados comuns.

#### <a name="common-data-types"></a>Tipo comuns de dados

Nos acessos a dados baseados em seleção e em associação, os conteúdos dos documentos são expostos por meio dos tipos de dados comuns a todos os aplicativos compatíveis do Office. No Office 2013, há suporte para três tipos de dados principais.

|**Tipo de dados**|**Descrição**|**Suporte ao aplicativo de host**|
|:-----|:-----|:-----|
|Texto|Fornece uma representação, em uma cadeia de caracteres, dos dados na seleção ou associação.|No Excel 2013, no Project 2013 e no PowerPoint 2013, há suporte apenas para texto sem formatação. No Word 2013, há suporte para três formatos de texto: texto sem formatação, HTML e OOXML (Office Open XML). Quando o texto é selecionado em uma célula no Excel, os métodos baseados em seleção realizam os processos de leitura e gravação para todo o conteúdo da célula, mesmo que apenas uma parte do texto esteja selecionada na célula. Quando texto é selecionado no Word e no PowerPoint, os métodos baseados em seleção realizam os processos de leitura e gravação apenas para os caracteres selecionados. O Project 2013 e o PowerPoint 2013 dão suporte apenas ao acesso a dados com base em seleção.|
|Matriz|Fornece os dados na seleção ou associação como uma **Array** bidimensional, que, no JavaScript, é implementada como uma matriz de matrizes. Por exemplo, duas linhas de valores **string** em duas colunas seriam ` [['a', 'b'], ['c', 'd']]`, e uma única coluna com três linhas seria `[['a'], ['b'], ['c']]`.|Há suporte ao acesso a dados de matriz apenas no Excel 2013 e no Word 2013.|
|Tabela|Fornece os dados na seleção ou associação como um objeto [TableData](/javascript/api/office/office.tabledata). O `TableData` objeto expõe os dados por meio das propriedades `headers` e `rows` .|Há suporte ao acesso a dados de tabela apenas no Excel 2013 e no Word 2013.|

#### <a name="data-type-coercion"></a>Coerção de tipo de dados

Os métodos de acesso `Document` a dados nos objetos [e Binding](/javascript/api/office/office.binding) suportam a especificação do tipo de dados desejado usando o parâmetro _coercionType_ desses métodos e os valores de enumeração [CoercionType](/javascript/api/office/office.coerciontype) correspondentes. Independentemente da forma real da associação, os diferentes aplicativos do Office dão suporte aos tipos de dados comuns ao tentar forçar os dados a usarem o tipo de dados solicitado. Por exemplo, se uma tabela ou um parágrafo do Word for selecionado, o desenvolvedor pode escolher se deseja lê-lo como texto sem formatação, Office Open XML ou tabela, e a implementação da API manipula as conversões de dados e as transformações necessárias.

> [!TIP]
> **Quando devo usar a matriz ou a tabela coercionType para o acesso aos dados?** Se você precisar que seus dados tabulares cresçam dinamicamente quando linhas e colunas são adicionadas e você deve trabalhar com os headers de tabela, você deve usar o tipo de dados de tabela (especificando o parâmetro _coercionType_ de um `Document` `Binding` ou método de acesso a dados de objeto como `"table"` ou `Office.CoercionType.Table`). A adição de linhas e colunas na estrutura de dados tem suporte nos dados de tabela e matriz, mas o acréscimo de linhas e colunas só tem suporte para dados de tabela. Se você não estiver planejando adicionar linhas e colunas e seus dados não exigirem funcionalidade de header, use o tipo de dados de matriz (especificando o parâmetro  _coercionType_ `"matrix"` do método de acesso a dados como ou `Office.CoercionType.Matrix`), o que fornece um modelo mais simples de interação com os dados.

Se os dados não puderem ser forçados para o tipo especificado, a propriedade [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) presente nos retornos de chamada retorna `"failed"`, e você pode usar a propriedade [AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) para acessar um objeto [Error](/javascript/api/office/office.error) com informações sobre o motivo pelo qual a chamada de método falhou.

## <a name="work-with-selections-using-the-document-object"></a>Trabalhar com seleções usando o objeto Document

O `Document` objeto expõe métodos que permitem que você leia e escreva para a seleção atual do usuário de forma "obter e definir". Para fazer isso, o `Document` objeto fornece os `getSelectedDataAsync` métodos e `setSelectedDataAsync` .

Para obter exemplos de códigos que demostram como realizar tarefas com seleções, consulte [Ler e gravar dados na seleção ativa em um documento ou uma planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).

## <a name="work-with-bindings-using-the-bindings-and-binding-objects"></a>Trabalhar com vinculações usando os objetos Bindings e Binding

O acesso a dados baseado em associação habilita os suplementos de conteúdo e painel de tarefas a acessarem de forma consistente determinada região de um documento ou uma planilha por meio de um identificador vinculado a uma associação. Primeiro, o suplemento precisa estabelecer a associação chamando um dos métodos que vinculam uma parte do documento a um identificador exclusivo: [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)), [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) ou [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)). Depois que a associação é estabelecida, o suplemento pode usar o identificador fornecido para acessar os dados contidos na região vinculada do documento ou da planilha. A criação de vinculações fornece o seguinte valor ao seu complemento.

- Permite o acesso a estruturas comuns de dados em aplicativos compatíveis do Office, como: tabelas, intervalos ou texto (uma execução contígua de caracteres).

- Habilita operações de leitura/gravação sem exigir que o usuário realize uma seleção.

- Estabelece uma relação entre o suplemento e os dados presentes no documento. As associações estão presentes no documento e podem ser acessadas em um momento posterior.

A criação de uma associação também permite que você se inscreva em eventos de alteração de seleção e de dados que apresentem um escopo definido para essa região específica do documento ou da planilha. Isso significa que o suplemento só é notificado sobre alterações que ocorrem dentro da região associada, e não sobre alterações gerais que ocorrem em todo o documento ou planilha.

O [objeto Bindings](/javascript/api/office/office.bindings) expõe um [método getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) que dá acesso ao conjunto de todas as vinculações estabelecidas no documento ou planilha. Uma associação individual pode ser acessada por sua ID usando os métodos [Bindings.getBindingByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) [ou Office.select](/javascript/api/office). Você pode estabelecer novas vinculações, bem como remover as existentes `Bindings` usando um dos seguintes métodos do objeto: [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)), [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)), [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)) ou [releaseByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-releasebyidasync-member(1)).

Há três tipos diferentes de vinculações que você especifica com o parâmetro  _bindingType_ `addFromSelectionAsync`ao criar uma associação com os métodos , `addFromPromptAsync` ou `addFromNamedItemAsync` .

|**Tipo de associação**|**Descrição**|**Suporte ao aplicativo de host**|
|:-----|:-----|:-----|
|Associação de texto|Associa a uma região do documento que pode ser representada como um texto.|No Word, a maioria das seleções contíguas são válidas, enquanto no Excel apenas as seleções de células únicas podem ser usadas para uma associação de texto. No Excel, só há suporte para texto sem formatação. No Word, há suporte para três formatos: texto sem formatação, HTML e Open XML do Office.|
|Associação de matriz|Associa a uma região fixa de um documento que contém dados tabulares sem cabeçalhos. Os dados de uma associação de matriz são gravados ou lidos como uma **Array** bidimensional, que é implementada como uma matriz de matrizes no JavaScript. Por exemplo, duas linhas de valores **string** em duas colunas podem ser gravadas ou lidas como ` [['a', 'b'], ['c', 'd']]`, e uma única coluna de três linhas pode ser gravada ou lida como `[['a'], ['b'], ['c']]`.|No Excel, qualquer seleção contígua de células pode ser usada para estabelecer uma associação de matriz. No Word, apenas as tabelas dão suporte à associação de matriz.|
|Associação de tabelas|Associa a uma região de um documento que contém uma tabela com cabeçalhos. Os dados em uma associação de tabela são gravados ou lidos como um objeto [TableData](/javascript/api/office/office.tabledata). O `TableData` objeto expõe os dados por meio das **propriedades de headers** **e linhas** .|Qualquer tabela do Excel ou Word pode ser a base para uma associação de tabela. Após estabelecer uma associação de tabelas, as linhas ou colunas novas que um usuário adicionar à tabela são automaticamente incluídas na associação.  |

<br/>

Depois que uma associação é criada usando um dos três métodos de "adicionar" `Bindings` do objeto, você pode trabalhar com os dados e propriedades da associação usando os métodos do objeto correspondente: [MatrixBinding](/javascript/api/office/office.matrixbinding), [TableBinding](/javascript/api/office/office.tablebinding) ou [TextBinding](/javascript/api/office/office.textbinding). Todos esses três objetos herdam os [métodos getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) e [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)) `Binding` do objeto que permitem que você interaja com os dados vinculados.

Para obter exemplos de códigos que demonstram como realizar tarefas com associações, consulte [Associar a regiões em um documento ou uma planilha](bind-to-regions-in-a-document-or-spreadsheet.md).

## <a name="work-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Trabalhar com partes XML personalizadas usando os objetos CustomXmlParts e CustomXmlPart

 **Aplica-se a:** suplementos de painel de tarefas para Word

Os objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts) e [CustomXmlPart](/javascript/api/office/office.customxmlpart) da API fornecem acesso a partes XML personalizadas de documentos do Word, que permitem a manipulação orientada por XML de conteúdo do documento. Para demonstrações de como trabalhar com `CustomXmlParts` os objetos e `CustomXmlPart` , consulte o exemplo de código [word-add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) .

## <a name="work-with-the-entire-document-using-the-getfileasync-method"></a>Trabalhar com o documento inteiro usando o método getFileAsync

 **Aplica-se a:** suplementos de painel de tarefas para Word e PowerPoint

O método [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) e os membros dos objetos [File](/javascript/api/office/office.file) e [Slice](/javascript/api/office/office.slice) fornecem a funcionalidade necessária para obter documentos inteiros do Word e PowerPoint em fatias (frações) de até 4 MB por vez. Para saber mais, consulte [Obter todo o documento por meio de um suplemento para PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).

## <a name="mailbox-object"></a>Objeto Mailbox

**Aplica-se a:** suplementos do Outlook

Os suplementos do Outlook usam principalmente um subconjunto da API exposta no objeto [Mailbox](/javascript/api/outlook/office.mailbox). Para acessar os objetos e membros específicos para suplementos do Outlook, como o objeto [Item](/javascript/api/outlook/office.item), use a propriedade [mailbox](/javascript/api/office/office.context#office-office-context-mailbox-member) do objeto **Context** para acessar o objeto **Mailbox**, conforme exibido na linha de código abaixo.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Além disso, Outlook os complementos podem usar os seguintes objetos.

- `Office` object: para inicialização.

- `Context` object: for access to content and display language properties.

- `RoamingSettings`object: para salvar Outlook configurações personalizadas específicas do complemento na caixa de correio do usuário onde o complemento está instalado.

Para obter informações sobre como usar o JavaScript em suplementos do Outlook, confira [Suplementos do Outlook ](../outlook/outlook-add-ins-overview.md).
