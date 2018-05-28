---
title: Associar a regi?es em um documento ou em uma planilha
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd26aa12e5d6da145fb6a2a89daf937cf6e88f04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a><span data-ttu-id="1014e-102">Associar a regi?es em um documento ou em uma planilha</span><span class="sxs-lookup"><span data-stu-id="1014e-102">Bind to regions in a document or spreadsheet</span></span>

<span data-ttu-id="1014e-p101">O acesso a dados baseado em associa??o permite que os suplementos de conte?do e de pain?is de tarefas acessem determinada regi?o de um documento ou planilha por meio de um identificador. Primeiro, o suplemento precisa estabelecer a associa??o. Para isso, ele chama um dos m?todos que associa uma parte do documento a um identificador exclusivo: [addFromPromptAsync], [addFromSelectionAsync] ou [addFromNamedItemAsync]. Depois que a associa??o ? estabelecida, o suplemento pode usar o identificador fornecido para acessar os dados contidos na regi?o associada do documento ou da planilha. A cria??o de associa??es proporciona o seguinte valor para o seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="1014e-p101">Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync], [addFromSelectionAsync], or [addFromNamedItemAsync]. After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in:</span></span>


- <span data-ttu-id="1014e-107">Permite o acesso a estruturas comuns de dados em aplicativos compat?veis do Office, como: tabelas, intervalos ou texto (uma execu??o cont?gua de caracteres).</span><span class="sxs-lookup"><span data-stu-id="1014e-107">Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).</span></span>
    
- <span data-ttu-id="1014e-108">Habilita opera??es de leitura/grava??o sem exigir que o usu?rio realize uma sele??o.</span><span class="sxs-lookup"><span data-stu-id="1014e-108">Enables read/write operations without requiring the user to make a selection.</span></span>
    
- <span data-ttu-id="1014e-p102">Estabelece uma rela??o entre o suplemento e os dados presentes no documento. As associa??es est?o presentes no documento e podem ser acessadas em um momento posterior.</span><span class="sxs-lookup"><span data-stu-id="1014e-p102">Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.</span></span>
    
<span data-ttu-id="1014e-p103">A cria??o de uma associa??o tamb?m permite que voc? se inscreva em eventos de altera??o de sele??o e de dados que apresentem um escopo definido para essa regi?o espec?fica do documento ou da planilha. Isso significa que o suplemento s? ? notificado sobre altera??es que ocorrem dentro da regi?o associada, e n?o sobre altera??es gerais que ocorrem em todo o documento ou planilha.</span><span class="sxs-lookup"><span data-stu-id="1014e-p103">Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.</span></span>

<span data-ttu-id="1014e-p104">O objeto [Bindings] exp?e um m?todo [getAllAsync], que d? acesso ao conjunto de todas as associa??es estabelecidas no documento ou na planilha. Uma associa??o individual pode ser acessada por sua ID, usando o m?todo Bindings.[getByIdAsync] ou [Office.select]. Voc? pode estabelecer novas associa??es e remover as existentes usando um dos seguintes m?todos do objeto [Bindings]: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync] ou [releaseByIdAsync].</span><span class="sxs-lookup"><span data-stu-id="1014e-p104">The [Bindings] object exposes a [getAllAsync] method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the Bindings.[getByIdAsync] or [Office.select] methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the [Bindings] object: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync], or [releaseByIdAsync].</span></span>


## <a name="binding-types"></a><span data-ttu-id="1014e-116">Tipos de associa??o</span><span class="sxs-lookup"><span data-stu-id="1014e-116">Binding types</span></span>

<span data-ttu-id="1014e-117">H? [tr?s tipos diferentes de associa??es][Office.BindingType] que podem ser especificadas com o par?metro _bindingType_ ao criar uma associa??o com os m?todos [addFromSelectionAsync], [addFromPromptAsync] ou [addFromNamedItemAsync]:</span><span class="sxs-lookup"><span data-stu-id="1014e-117">There are [three different types of bindings][Office.BindingType] that you specify with the  _bindingType_ parameter when you create a binding with the [addFromSelectionAsync], [addFromPromptAsync] or [addFromNamedItemAsync] methods:</span></span>

1. <span data-ttu-id="1014e-118">**[Text Binding][TextBinding]**: associa a uma regi?o do documento que pode ser representada como texto.</span><span class="sxs-lookup"><span data-stu-id="1014e-118">**[Text Binding][TextBinding]** - Binds to a region of the document that can be represented as text.</span></span>

    <span data-ttu-id="1014e-p105">No Word, a maioria das sele??es cont?guas s?o v?lidas, enquanto no Excel apenas as sele??es de c?lulas ?nicas podem ser usadas para uma associa??o de texto. No Excel, s? h? suporte para texto sem formata??o. No Word, h? suporte para tr?s formatos: texto sem formata??o, HTML e Open XML do Office.</span><span class="sxs-lookup"><span data-stu-id="1014e-p105">In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.</span></span>

2. <span data-ttu-id="1014e-p106">**[Matrix Binding][MatrixBinding]**: associa uma regi?o fixa de um documento que cont?m dados tabulares sem cabe?alhos. Os dados em uma associa??o de matriz s?o gravados ou lidos como uma **Array** bidimensional, que ? implementada no JavaScript como uma matriz de matrizes. Por exemplo, duas linhas de valores da  **cadeia de caracteres** em duas colunas podem ser gravadas ou lidas como ` [['a', 'b'], ['c', 'd']]` e uma ?nica coluna de tr?s linhas pode ser gravada ou lida como  `[['a'], ['b'], ['c']]`.</span><span class="sxs-lookup"><span data-stu-id="1014e-p106">**[Matrix Binding][MatrixBinding]** - Binds to a fixed region of a document that contains tabular data without headers.Data in a matrix binding is written or read as a two dimensional  **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of  **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as  `[['a'], ['b'], ['c']]`.</span></span>

    <span data-ttu-id="1014e-p107">No Excel, qualquer sele??o cont?gua de c?lulas pode ser usada para estabelecer uma associa??o de matriz. No Word, apenas as tabelas d?o suporte ? associa??o de matriz.</span><span class="sxs-lookup"><span data-stu-id="1014e-p107">In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.</span></span>

3. <span data-ttu-id="1014e-p108">**[Table Binding][TableBinding]**: associa uma regi?o de um documento que cont?m uma tabela com cabe?alhos. Os dados em uma associa??o de tabela s?o gravados ou lidos como um objeto [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). O objeto `TableData` exp?e os dados por meio das propriedades `headers` e `rows`.</span><span class="sxs-lookup"><span data-stu-id="1014e-p108">**[Table Binding][TableBinding]** - Binds to a region of a document that contains a table with headers.Data in a table binding is written or read as a [TableData](https://dev.office.com/reference/add-ins/shared/tabledata) object. The `TableData` object exposes the data through the `headers` and `rows` properties.</span></span>

    <span data-ttu-id="1014e-p109">Qualquer tabela do Excel ou Word pode ser a base para uma associa??o de tabela. Ap?s estabelecer uma associa??o de tabelas, as linhas ou colunas novas que um usu?rio adicionar ? tabela s?o automaticamente inclu?das na associa??o.</span><span class="sxs-lookup"><span data-stu-id="1014e-p109">Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding.</span></span>

<span data-ttu-id="1014e-p110">Depois que uma associa??o ? criada usando um dos tr?s m?todos "addFrom" do objeto `Bindings` ? poss?vel trabalhar com dados e as propriedades da associa??o usando os m?todos do objeto correspondente: [MatrixBinding], [TableBinding] ou [TextBinding]. Esses tr?s objetos herdam os m?todos  [getDataAsync] e [setDataAsync] do objeto `Binding`, o que permite interagir com os dados associados.</span><span class="sxs-lookup"><span data-stu-id="1014e-p110">After a binding is created by using one of the three "addFrom" methods of the  `Bindings` object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding], [TableBinding], or [TextBinding]. All three of these objects inherit the [getDataAsync] and [setDataAsync] methods of the `Binding` object that enable you to interact with the bound data.</span></span>

> [!NOTE]
> <span data-ttu-id="1014e-p111">**Quando devo usar a matriz ou as associa??es de tabela?** Quando os dados tabulares com os quais voc? est? trabalhando contiverem uma linha de totais, voc? dever? usar uma associa??o de matriz se o script do suplemento precisar acessar valores na linha de totais ou detectar que a sele??o do usu?rio est? na linha de totais. Se voc? estabelecer uma associa??o de tabela para os dados tabulares que cont?m uma linha de totais, a propriedade [TableBinding.rowCount] e as propriedades `rowCount` and `startRow` do objeto [BindingSelectionChangedEventArgs] nos manipuladores de eventos n?o refletir?o a linha de totais em seus valores. Para resolver essa limita??o, voc? deve estabelecer uma associa??o de matriz para trabalhar com a linha de totais.</span><span class="sxs-lookup"><span data-stu-id="1014e-p111">**When should you use matrix versus table bindings?** When the tabular data you are working with contains a total row, you must use a matrix binding if your add-in's script needs to access values in the total row or detect that the user's selection is in the total row. If you establish a table binding for tabular data that contains a total row, the [TableBinding.rowCount] property and the `rowCount` and `startRow` properties of the [BindingSelectionChangedEventArgs] object in event handlers won't reflect the total row in their values. To work around this limitation, you must use establish a matrix binding to work with the total row.</span></span>

## <a name="add-a-binding-to-the-users-current-selection"></a><span data-ttu-id="1014e-136">Adicionar uma associa??o ? sele??o atual do usu?rio</span><span class="sxs-lookup"><span data-stu-id="1014e-136">Add a binding to the user's current selection</span></span>

<span data-ttu-id="1014e-137">O exemplo a seguir mostra como adicionar uma associa??o de texto chamada `myBinding` ? sele??o atual em um documento usando o m?todo [addFromSelectionAsync].</span><span class="sxs-lookup"><span data-stu-id="1014e-137">The following example shows how to add a text binding called  `myBinding` to the current selection in a document by using the [addFromSelectionAsync] method.</span></span>


```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="1014e-p112">Neste exemplo, o tipo de associa??o especificado ? texto. Isso significa que um [TextBinding] ser? criado para a sele??o. Diferentes tipos de associa??o exp?em dados e opera??es diferentes. [Office.BindingType] ? uma enumera??o de valores de tipos de associa??es dispon?veis.</span><span class="sxs-lookup"><span data-stu-id="1014e-p112">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection. Different binding types expose different data and operations. [Office.BindingType] is an enumeration of available binding type values.</span></span>

<span data-ttu-id="1014e-p113">O segundo par?metro opcional ? um objeto que especifica a ID da nova associa??o que est? sendo criada. Se uma ID n?o for especificada, uma ser? gerada automaticamente.</span><span class="sxs-lookup"><span data-stu-id="1014e-p113">The second optional parameter is an object that specifies the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="1014e-p114">A fun??o an?nima transmitida para a fun??o como o par?metro _callback_ final ? executada quando a cria??o da associa??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, `asyncResult`, que fornece acesso a um objeto [AsyncResult] que fornece o status da chamada. A propriedade `AsyncResult.value` cont?m uma refer?ncia para um objeto [Binding] do tipo especificado para a associa??o rec?m-criada. Voc? pode usar esse objeto [Binding] para obter e definir os dados.</span><span class="sxs-lookup"><span data-stu-id="1014e-p114">The anonymous function that is passed into the function as the final  _callback_ parameter is executed when the creation of the binding is complete. The function is called with a single parameter, `asyncResult`, which provides access to an [AsyncResult] object that provides the status of the call. The `AsyncResult.value` property contains a reference to a [Binding] object of the type that is specified for the newly created binding. You can use this [Binding] object to get and set data.</span></span>

## <a name="add-a-binding-from-a-prompt"></a><span data-ttu-id="1014e-148">Adicionar uma associa??o a partir de um prompt</span><span class="sxs-lookup"><span data-stu-id="1014e-148">Add a binding from a prompt</span></span>

<span data-ttu-id="1014e-p115">O exemplo a seguir mostra como adicionar uma associa??o de texto chamada `myBinding` usando o m?todo [addFromPromptAsync]. Este m?todo permite ao usu?rio especificar o intervalo da associa??o usando o prompt de sele??o de intervalo interno do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="1014e-p115">The following example shows how to add a text binding called  `myBinding` by using the [addFromPromptAsync] method. This method lets the user specify the range for the binding by using the application's built-in range selection prompt.</span></span>


```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="1014e-p116">Neste exemplo, o tipo de associa??o especificado ? texto. Isso significa que um [TextBinding] ser? criado para a sele??o que o usu?rio especificar no prompt.</span><span class="sxs-lookup"><span data-stu-id="1014e-p116">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection that the user specifies in the prompt.</span></span>

<span data-ttu-id="1014e-p117">O segundo par?metro ? um objeto que cont?m a ID da nova associa??o que est? sendo criada. Se uma ID n?o for especificada, uma ser? gerada automaticamente.</span><span class="sxs-lookup"><span data-stu-id="1014e-p117">The second parameter is an object that contains the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="1014e-p118">A fun??o an?nima transmitida para a fun??o como o terceiro par?metro _callback_ ? executada quando a cria??o da associa??o ? conclu?da. Quando a fun??o de retorno de chamada ? executada, o objeto [AsyncResult] cont?m o status da chamada e a associa??o rec?m-criada.</span><span class="sxs-lookup"><span data-stu-id="1014e-p118">The anonymous function passed into the function as the third  _callback_ parameter is executed when the creation of the binding is complete. When the callback function executes, the [AsyncResult] object contains the status of the call and the newly created binding.</span></span>

<span data-ttu-id="1014e-157">A Figura 1 mostra o prompt de sele??o do intervalo interno no Excel.</span><span class="sxs-lookup"><span data-stu-id="1014e-157">Figure 1 shows the built-in range selection prompt in Excel.</span></span>


<span data-ttu-id="1014e-158">*Figura 1. Selecionar IU de Dados do Excel*</span><span class="sxs-lookup"><span data-stu-id="1014e-158">*Figure 1. Excel Select Data UI*</span></span>

![Selecionar IU de Dados do Excel](../images/agave-api-overview-excel-selection-ui.png)


## <a name="add-a-binding-to-a-named-item"></a><span data-ttu-id="1014e-160">Adicionar uma associa??o a um item nomeado</span><span class="sxs-lookup"><span data-stu-id="1014e-160">Add a binding to a named item</span></span>


<span data-ttu-id="1014e-161">O exemplo a seguir mostra como adicionar uma associa??o ao item nomeado `myRange` existente como uma associa??o de "matriz" usando o m?todo [addFromNamedItemAsync] e atribui a `id` da associa??o como "myMatrix".</span><span class="sxs-lookup"><span data-stu-id="1014e-161">The following example shows how to add a binding to the existing  `myRange` named item as a "matrix" binding by using the [addFromNamedItemAsync] method, and assigns the binding's `id` as "myMatrix".</span></span>


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

<span data-ttu-id="1014e-p119">**Para Excel**, o par?metro `itemName` do m?todo [addFromNamedItemAsync] pode se referir a um intervalo nomeado existente, a um intervalo especificado com o estilo de refer?ncia `A1` `("A1:A3")` ou a uma tabela. Por padr?o, adicionar uma tabela no Excel atribui o nome "Tabela1" ? primeira tabela adicionada, "Tabela2" ? segunda tabela adicionada e assim por diante. Para atribuir um nome significativo para uma tabela na IU do Excel, use a propriedade **Table Name** na guia **Ferramentas da Tabela | Design** da faixa de op??es.</span><span class="sxs-lookup"><span data-stu-id="1014e-p119">**For Excel**, the  `itemName` parameter of the [addFromNamedItemAsync] method can refer to an existing named range, a range specified with the `A1` reference style `("A1:A3")`, or a table. By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on. To assign a meaningful name for a table in the Excel UI, use the **Table Name** property on the **Table Tools | Design** tab of the ribbon.</span></span>


> [!NOTE]
> <span data-ttu-id="1014e-165">No Excel, ao especificar uma tabela como um item nomeado, ? preciso qualificar totalmente o nome ao incluir o nome da planilha no nome da tabela neste formato: `"Sheet1!Table1"`.  `"Sheet1!Table1"`</span><span class="sxs-lookup"><span data-stu-id="1014e-165">In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of the table in this format:  `"Sheet1!Table1"`</span></span>

<span data-ttu-id="1014e-166">O exemplo a seguir cria uma associa??o no Excel para as tr?s primeiras c?lulas na coluna A (`"A1:A3"`), atribui a ID `"MyCities"` e, em seguida, grava tr?s nomes de cidades ? associa??o.</span><span class="sxs-lookup"><span data-stu-id="1014e-166">The following example creates a binding in Excel to the first three cells in column A ( `"A1:A3"`), assigns the  id `"MyCities"`, and then writes three city names to that binding.</span></span>


```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="1014e-p120">**Para o Word**, o par?metro `itemName` do m?todo [addFromNamedItemAsync] refere-se ? propriedade `Title` de um controle de conte?do `Rich Text`. (N?o ? poss?vel associar controles de conte?do diferentes do controle de conte?do `Rich Text`.)</span><span class="sxs-lookup"><span data-stu-id="1014e-p120">**For Word**, the  `itemName` parameter of the [addFromNamedItemAsync] method refers to the `Title` property of a `Rich Text` content control. (You can't bind to content controls other than the `Rich Text` content control.)</span></span>

<span data-ttu-id="1014e-p121">Por padr?o, um controle de conte?do n?o tem um valor `Title*` atribu?do. Para atribuir um nome significativo na IU do Word, ap?s inserir um controle de conte?do **Rich Text** do grupo **Controles** na guia **Desenvolvedor** da faixa de op??es, use o comando **Propriedades** no grupo **Controles** para exibir a caixa de di?logo **Propriedades de Controle do Conte?do**. Em seguida, defina a propriedade **Title** do controle de conte?do para o nome que voc? deseja referenciar a partir de seu c?digo.</span><span class="sxs-lookup"><span data-stu-id="1014e-p121">By default, a content control has no  `Title*`value assigned. To assign a meaningful name in the Word UI, after inserting a **Rich Text** content control from the **Controls** group on the **Developer** tab of the ribbon, use the **Properties** command in the **Controls** group to display the **Content Control Properties** dialog box. Then set the **Title** property of the content control to the name you want to reference from your code.</span></span>

<span data-ttu-id="1014e-172">O exemplo a seguir cria uma associa??o de texto no Word para um controle de conte?do de rich text denominado `"FirstName"`, atribui a **id**`"firstName"` e, em seguida, exibe essas informa??es.</span><span class="sxs-lookup"><span data-stu-id="1014e-172">The following example creates a text binding in Word to a rich text content control named  `"FirstName"`, assigns the  **id** `"firstName"`, and then displays that information.</span></span>


```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="get-all-bindings"></a><span data-ttu-id="1014e-173">Obter todas as associa??es</span><span class="sxs-lookup"><span data-stu-id="1014e-173">Get all bindings</span></span>


<span data-ttu-id="1014e-174">O exemplo a seguir mostra como obter todas as associa??es em um documento usando o m?todo Bindings.[getAllAsync].</span><span class="sxs-lookup"><span data-stu-id="1014e-174">The following example shows how to get all bindings in a document by using the Bindings.[getAllAsync] method.</span></span>


```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="1014e-p122">A fun??o an?nima transmitida para a fun??o como o par?metro `callback` ? executada quando a opera??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, `asyncResult`, que cont?m uma matriz das associa??es no documento. A matriz ? repetida para compilar uma cadeia de caracteres contendo as IDs das associa??es. A cadeia de caracteres ?, ent?o, exibida em uma caixa de mensagem.</span><span class="sxs-lookup"><span data-stu-id="1014e-p122">The anonymous function that is passed into the function as the  `callback` parameter is executed when the operation is complete. The function is called with a single parameter, `asyncResult`, which contains an  array of the bindings in the document. The array is iterated to build a string that contains the IDs of the bindings. The string is then displayed in a message box.</span></span>


## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a><span data-ttu-id="1014e-179">Obter uma associa??o por ID usando o m?todo getByIdAsync do objeto Bindings</span><span class="sxs-lookup"><span data-stu-id="1014e-179">Get a binding by ID using the getByIdAsync method of the Bindings object</span></span>


<span data-ttu-id="1014e-p123">O exemplo a seguir mostra como usar o m?todo [getByIdAsync] para obter uma associa??o em um documento ao especificar sua ID. Este exemplo sup?e que uma associa??o nomeada `'myBinding'` foi adicionada ao documento usando um dos m?todos descritos anteriormente neste t?pico.</span><span class="sxs-lookup"><span data-stu-id="1014e-p123">The following example shows how to use the [getByIdAsync] method to get a binding in a document by specifying its ID. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>


```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } 
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="1014e-182">No exemplo, o primeiro par?metro `id` ? a ID da associa??o a recuperar.</span><span class="sxs-lookup"><span data-stu-id="1014e-182">In the example, the first  `id` parameter is the ID of the binding to retrieve.</span></span>

<span data-ttu-id="1014e-p124">A fun??o an?nima que ? transmitida para a fun??o como o segundo par?metro _callback_ ? executada quando a opera??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, _asyncResult_, que cont?m o status da chamada e a associa??o com a ID "myBinding".</span><span class="sxs-lookup"><span data-stu-id="1014e-p124">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the operation is completed. The function is called with a single parameter, _asyncResult_, which contains the status of the call and the binding with the ID "myBinding".</span></span>


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a><span data-ttu-id="1014e-185">Obter uma associa??o pela ID usando o m?todo selecionado do objeto Office</span><span class="sxs-lookup"><span data-stu-id="1014e-185">Get a binding by ID using the select method of the Office object</span></span>


<span data-ttu-id="1014e-p125">O exemplo a seguir mostra como usar o m?todo [Office.select] para obter a promessa de um objeto [Binding] em um documento especificando sua ID em uma cadeia de caracteres do seletor. Em seguida, chama o m?todo Binding.[getDataAsync] para obter os dados na associa??o especificada. Este exemplo sup?e que uma associa??o denominada `'myBinding'` foi adicionada ao documento usando um dos m?todos descritos anteriormente neste t?pico.</span><span class="sxs-lookup"><span data-stu-id="1014e-p125">The following example shows how to use the [Office.select] method to get a [Binding] object promise in a document by specifying its ID in a selector string. It then calls the Binding.[getDataAsync] method to get data from the specified binding. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>


```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


> [!NOTE]
> <span data-ttu-id="1014e-p126">Se a promessa do m?todo `select` retornar um objeto [Binding] com ?xito, esse objeto ir? expor somente os seguintes quatro m?todos do objeto: [getDataAsync], [setDataAsync], [addHandlerAsync], e [removeHandlerAsync]. Se a promessa n?o puder retornar um objeto Binding, o retorno de chamada `onError` pode ser usado para acessar um objeto [asyncResult].error para mais informa??es. Se for preciso chamar um membro do objeto Binding diferente dos quatro m?todos expostos pela promessa do objeto Binding retornada pelo m?todo `select`, use o m?todo [getByIdAsync] utilizando a propriedade [Document.bindings] e o m?todo Bindings.[getByIdAsync] para recuperar o objeto Binding**.</span><span class="sxs-lookup"><span data-stu-id="1014e-p126">If the  `select` method promise successfully returns a [Binding] object, that object exposes only the following four methods of the object: [getDataAsync], [setDataAsync], [addHandlerAsync], and [removeHandlerAsync]. If the promise cannot return a  Binding object, the `onError` callback can be used to access an [asyncResult].error object to get more information.If you need to call a member of the Binding object other than the four methods exposed by the Binding object promise returned by the `select` method, instead use the [getByIdAsync] method by using the [Document.bindings] property and Bindings.[getByIdAsync] method to retrieve the Binding** object.</span></span>

## <a name="release-a-binding-by-id"></a><span data-ttu-id="1014e-191">Liberar uma associa??o pela ID</span><span class="sxs-lookup"><span data-stu-id="1014e-191">Release a binding by ID</span></span>


<span data-ttu-id="1014e-192">O exemplo a seguir mostra como usar o m?todo [releaseByIdAsync] para liberar uma associa??o em um documento, especificando sua ID.</span><span class="sxs-lookup"><span data-stu-id="1014e-192">The following example shows how use the [releaseByIdAsync] method to release a binding in a document by specifying its ID.</span></span>

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="1014e-193">No exemplo, o primeiro par?metro `id` ? a ID da associa??o a liberar.</span><span class="sxs-lookup"><span data-stu-id="1014e-193">In the example, the first `id` parameter is the ID of the binding to release.</span></span>

<span data-ttu-id="1014e-p127">A fun??o an?nima que ? transmitida para a fun??o como o segundo par?metro ? um retorno de chamada executado quando a opera??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, [asyncResult], que cont?m o status da chamada.</span><span class="sxs-lookup"><span data-stu-id="1014e-p127">The anonymous function that is passed into the function as the second parameter is a callback that is executed when the operation is complete. The function is called with a single parameter,  [asyncResult], which contains the status of the call.</span></span>


## <a name="read-data-from-a-binding"></a><span data-ttu-id="1014e-196">Ler os dados de uma associa??o</span><span class="sxs-lookup"><span data-stu-id="1014e-196">Read data from a binding</span></span>


<span data-ttu-id="1014e-197">O exemplo a seguir mostra como usar o m?todo [getDataAsync] para obter dados de uma associa??o existente.</span><span class="sxs-lookup"><span data-stu-id="1014e-197">The following example shows how to use the [getDataAsync] method to get data from an existing binding.</span></span>


```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="1014e-p128">`myBinding` ? uma vari?vel que cont?m uma associa??o de texto existente no documento. Como alternativa, ? poss?vel usar [Office.select] para acessar a associa??o pela ID, e iniciar sua chamada para o m?todo [getDataAsync], assim:</span><span class="sxs-lookup"><span data-stu-id="1014e-p128">`myBinding` is a variable that contains an existing text binding in the document. Alternatively, you could use the [Office.select] to access the binding by its ID, and start your call to the [getDataAsync] method, like this:</span></span> 

```js 
Office.select("bindings#myBindingID").getDataAsync
```


<span data-ttu-id="1014e-p129">A fun??o an?nima transmitida para a fun??o ? um retorno de chamada executado quando a opera??o ? conclu?da. A propriedade [AsyncResult].value cont?m os dados em `myBinding`. O tipo do valor depende do tipo de associa??o. A associa??o neste exemplo ? uma associa??o de texto. Portanto, o valor conter? uma cadeia de caracteres. Para obter mais exemplos de como trabalhar com as associa??es de tabela e matriz, confira o t?pico do m?todo [getDataAsync].</span><span class="sxs-lookup"><span data-stu-id="1014e-p129">The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The [AsyncResult].value property contains the data within `myBinding`. The type of the value depends on the binding type. The binding in this example is a text binding. Therefore, the value will contain a string. For additional examples of working with matrix and table bindings, see the [getDataAsync] method topic.</span></span>


## <a name="write-data-to-a-binding"></a><span data-ttu-id="1014e-206">Gravar dados em uma associa??o</span><span class="sxs-lookup"><span data-stu-id="1014e-206">Write data to a binding</span></span>

<span data-ttu-id="1014e-207">O exemplo a seguir mostra como usar o m?todo [setDataAsync] para definir os dados em uma associa??o existente.</span><span class="sxs-lookup"><span data-stu-id="1014e-207">The following example shows how to use the [setDataAsync] method to set data in an existing binding.</span></span>

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 <span data-ttu-id="1014e-208">`myBinding` ? uma vari?vel que cont?m uma associa??o de texto existente no documento.</span><span class="sxs-lookup"><span data-stu-id="1014e-208">`myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="1014e-p130">No exemplo, o primeiro par?metro ? o valor a definir em `myBinding`. Como esta ? uma associa??o de texto, o valor ? uma `string`. Diferentes tipos de associa??o aceitam diferentes tipos de dados.</span><span class="sxs-lookup"><span data-stu-id="1014e-p130">In the example, the first parameter is the value to set on  `myBinding`. Because this is a text binding, the value is a `string`. Different binding types accept different types of data.</span></span>

<span data-ttu-id="1014e-p131">A fun??o an?nima que ? transmitida para a fun??o ? um retorno de chamada executado quando a opera??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, `asyncResult`, que cont?m o status do resultado.</span><span class="sxs-lookup"><span data-stu-id="1014e-p131">The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The function is called with a single parameter,  `asyncResult`, which contains the status of the result.</span></span>

> [!NOTE]
> <span data-ttu-id="1014e-214">A partir da vers?o do Excel 2013 SP1 e da compila??o correspondente do Excel Online, agora ? poss?vel [definir a formata??o ao escrever e atualizar dados em tabelas de vincula??o](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="1014e-214">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing and updating data in bound tables](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a><span data-ttu-id="1014e-215">Detectar altera??es nos dados ou a sele??o em uma associa??o</span><span class="sxs-lookup"><span data-stu-id="1014e-215">Detect changes to data or the selection in a binding</span></span>


<span data-ttu-id="1014e-216">O exemplo a seguir mostra como anexar um manipulador de eventos ao evento [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) de uma associa??o com uma id "MyBinding".</span><span class="sxs-lookup"><span data-stu-id="1014e-216">The following example shows how to attach an event handler to the [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event of a binding with an id of "MyBinding".</span></span>


```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="1014e-217">? uma vari?vel que cont?m uma associa??o de texto existente no documento.`myBinding`</span><span class="sxs-lookup"><span data-stu-id="1014e-217">The `myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="1014e-p132">O primeiro par?metro `eventType` do m?todo [addHandlerAsync] especifica o nome do evento no qual se inscrever. [Office.EventType] ? uma enumera??o dos valores do tipo de evento dispon?veis. `Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"\`.</span><span class="sxs-lookup"><span data-stu-id="1014e-p132">The first  `eventType` parameter of the [addHandlerAsync] method specifies the name of the event to subscribe to. [Office.EventType] is an enumeration of available event type values. `Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"\`.</span></span>

<span data-ttu-id="1014e-p133">A fun??o `dataChanged` que ? transmitida para a fun??o como o segundo par?metro _handler_ ? um manipulador de eventos executado quando os dados na associa??o s?o alterados. A fun??o ? chamada com um ?nico par?metro, _eventArgs_, que cont?m uma refer?ncia para a associa??o. Essa associa??o pode ser usada para recuperar os dados atualizados.</span><span class="sxs-lookup"><span data-stu-id="1014e-p133">The  `dataChanged` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the data in the binding is changed. The function is called with a single parameter, _eventArgs_, which contains a reference to the binding. This binding can be used to retrieve the updated data.</span></span>

<span data-ttu-id="1014e-p134">Da mesma forma, ? poss?vel detectar quando um usu?rio altera a sele??o em uma associa??o anexando um manipulador de eventos ao evento [SelectionChanged] de uma associa??o. Para fazer isso, especifique o par?metro `eventType` do m?todo [addHandlerAsync] como `Office.EventType.BindingSelectionChanged` ou `"bindingSelectionChanged"`.</span><span class="sxs-lookup"><span data-stu-id="1014e-p134">Similarly, you can detect when a user changes selection in a binding by attaching an event handler to the [SelectionChanged] event of a binding. To do that, specify the `eventType` parameter of the [addHandlerAsync] method as `Office.EventType.BindingSelectionChanged` or `"bindingSelectionChanged"`.</span></span>

<span data-ttu-id="1014e-p135">Voc? pode adicionar v?rios manipuladores de eventos para um determinado evento chamando o m?todo [addHandlerAsync] novamente e transmitindo uma fun??o do manipulador de eventos adicional para o par?metro `handler`. Isso funcionar? corretamente, contanto que o nome de cada fun??o do manipulador de eventos seja exclusivo.</span><span class="sxs-lookup"><span data-stu-id="1014e-p135">You can add multiple event handlers for a given event by calling the [addHandlerAsync] method again and passing in an additional event handler function for the `handler` parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


### <a name="remove-an-event-handler"></a><span data-ttu-id="1014e-228">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="1014e-228">Remove an event handler</span></span>


<span data-ttu-id="1014e-p136">Para remover um manipulador de eventos de um evento, chame o m?todo [removeHandlerAsync] passando o tipo de evento como o primeiro par?metro _eventType_ e o nome da fun??o do manipulador de eventos a remover como o segundo par?metro _handler_. Por exemplo, a fun??o a seguir remover? a fun??o de manipulador de eventos `dataChanged` adicionada no exemplo da se??o anterior.</span><span class="sxs-lookup"><span data-stu-id="1014e-p136">To remove an event handler for an event, call the [removeHandlerAsync] method passing in the event type as the first _eventType_ parameter, and the name of the event handler function to remove as the second _handler_ parameter. For example, the following function will remove the `dataChanged` event handler function added in the previous section's example.</span></span>


```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


> [!IMPORTANT]
> <span data-ttu-id="1014e-231">Se o par?metro opcional _handler_ for omitido ao chamar o m?todo [removeHandlerAsync], todos os manipuladores de eventos do `eventType` especificado ser?o removidos.</span><span class="sxs-lookup"><span data-stu-id="1014e-231">If the optional  _handler_ parameter is omitted when the [removeHandlerAsync] method is called, all event handlers for the specified `eventType` will be removed.</span></span>


## <a name="see-also"></a><span data-ttu-id="1014e-232">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="1014e-232">See also</span></span>

- [<span data-ttu-id="1014e-233">No??es b?sicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="1014e-233">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md) 
- [<span data-ttu-id="1014e-234">Programa??o ass?ncrona nos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1014e-234">Asynchronous programming in Office Add-ins</span></span>](asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="1014e-235">Leia e grave dados na sele??o ativa, em um documento ou em uma planilha</span><span class="sxs-lookup"><span data-stu-id="1014e-235">Read and write data to the active selection in a document or spreadsheet</span></span>](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
[Associa??o]:               https://dev.office.com/reference/add-ins/shared/binding
[Binding]:               https://dev.office.com/reference/add-ins/shared/binding
[MatrixBinding]:         https://dev.office.com/reference/add-ins/shared/binding.matrixbinding
[TableBinding]:          https://dev.office.com/reference/add-ins/shared/binding.tablebinding
[TextBinding]:           https://dev.office.com/reference/add-ins/shared/binding.textbinding
[getDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.getdataasync
[setDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.setdataasync
[SelectionChanged]:      https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent
[addHandlerAsync]:       https://dev.office.com/reference/add-ins/shared/binding.addhandlerasync
[removeHandlerAsync]:    https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync

[Associa??es]:              https://dev.office.com/reference/add-ins/shared/bindings.bindings
[Bindings]:              https://dev.office.com/reference/add-ins/shared/bindings.bindings
[getByIdAsync]:          https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync 
[getAllAsync]:           https://dev.office.com/reference/add-ins/shared/bindings.getallasync
[addFromNamedItemAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync
[addFromSelectionAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync
[addFromPromptAsync]:    https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync
[releaseByIdAsync]:      https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync

[AsyncResult]:          https://dev.office.com/reference/add-ins/shared/asyncresult
[Office.BindingType]:   https://dev.office.com/reference/add-ins/shared/bindingtype-enumeration
[Office.select]:        https://dev.office.com/reference/add-ins/shared/office.select 
[Office.EventType]:     https://dev.office.com/reference/add-ins/shared/eventtype-enumeration 
[Document.bindings]:    https://dev.office.com/reference/add-ins/shared/document.bindings


[TableBinding.rowCount]: https://dev.office.com/reference/add-ins/shared/binding.tablebinding.rowcount
[BindingSelectionChangedEventArgs]: https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedeventargs
