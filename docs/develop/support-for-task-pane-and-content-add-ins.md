---
title: Suporte da API JavaScript para Office para suplementos de conteúdo e de painel de tarefas no Office 2013
description: Use a API JavaScript do Office para criar um painel de tarefas no Office 2013.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 185397073a9d62a868b5ebc065d94c57b778e1e7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718794"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Suporte da API JavaScript para Office para suplementos de conteúdo e de painel de tarefas no Office 2013

[!include[information about the common API](../includes/alert-common-api-info.md)]

Você pode usar a [API JavaScript para Office](../reference/javascript-api-for-office.md) para criar suplementos de painel de tarefas ou de conteúdo para aplicativos host do Office 2013. Os objetos e métodos que dão suporte a suplementos de conteúdo e de painel de tarefas são categorizados da seguinte forma:

1. **Objetos comuns compartilhados com outros suplementos do Office.** Esses objetos incluem [Office](/javascript/api/office), [Context](/javascript/api/office/office.context)e [AsyncResult](/javascript/api/office/office.asyncresult). O `Office` objeto é o objeto raiz da API JavaScript do Office. O `Context` objeto representa o ambiente de tempo de execução do suplemento. Ambos `Office` e `Context` são os objetos fundamentais para qualquer suplemento do Office. O `AsyncResult` objeto representa os resultados de uma operação assíncrona, como os dados retornados para o `getSelectedDataAsync` método, que lê o que um usuário selecionou em um documento.

2. **O objeto Document.** A maior parte da API disponível para os suplementos de conteúdo e de painel de tarefas é exposta por meio dos métodos, propriedades e eventos do objeto [Document](/javascript/api/office/office.document) . Um suplemento de conteúdo ou de painel de tarefas pode usar a propriedade [Office. Context. Document](/javascript/api/office/office.context#document) para acessar o objeto **Document** e por meio dele, pode acessar os principais membros da API para trabalhar com dados em documentos, como [associações](/javascript/api/office/office.bindings) e objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts) , e os métodos [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)e [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) . O `Document` objeto também fornece a propriedade [Mode](/javascript/api/office/office.document#mode) para determinar se um documento é somente leitura ou no modo de edição, a propriedade [URL](/javascript/api/office/office.document#url) para obter a URL do documento atual e o acesso ao objeto [Settings](/javascript/api/office/office.settings) . O `Document` objeto também oferece suporte à adição de manipuladores de eventos para o evento [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) , para que você possa detectar quando um usuário altera sua seleção no documento.

   Um suplemento de conteúdo ou de painel de tarefas pode acessar `Document` o objeto somente depois que o dom e o ambiente de tempo de execução foram carregados, normalmente no manipulador de eventos do evento [Office. Initialize](/javascript/api/office) . Para saber mais sobre o fluxo de eventos quando um suplemento é inicializado e como verificar se o DOM e o tempo de execução foram carregados com êxito, confira [Carregar o DOM e o ambiente de tempo de execução](loading-the-dom-and-runtime-environment.md).

3. **Objetos para trabalhar com recursos específicos.** Para trabalhar com recursos específicos da API, use as seguintes objetos e métodos:

    - Os métodos do objeto [Bindings](/javascript/api/office/office.bindings) para criar ou obter associações e os métodos e propriedades do objeto [Binding](/javascript/api/office/office.binding) para trabalhar com dados.

    - Os objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) e objetos associados para criar e manipular partes XML personalizadas em documentos do Word.

    - Os objetos [File](/javascript/api/office/office.file) e [Slice](/javascript/api/office/office.slice) para criar uma cópia do documento inteiro, dividi-lo em partes ou "fatias" e ler ou transmitir os dados nessas fatias.

    - O objeto [Settings](/javascript/api/office/office.settings) para salvar dados personalizados, como preferências do usuário e o estado do suplemento.


> [!IMPORTANT]
> Alguns membros da API não têm suporte em todos os aplicativos do Office que podem hospedar suplementos de conteúdo e de painel de tarefas. Para determinar quais membros têm suporte, confira o seguinte:

Para obter um resumo do suporte à API JavaScript do Office em aplicativos host do Office, consulte [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Ler e gravar em uma seleção ativa

Você pode ler ou gravar na seleção atual do usuário em um documento, planilha ou apresentação. Dependendo do aplicativo host para o suplemento, você pode especificar o tipo de estrutura de dados para ler ou gravar como um parâmetro nos métodos [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) e [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) do objeto [Document](/javascript/api/office/office.document). Por exemplo, você pode especificar qualquer tipo de dados (texto, HTML, dados tabulares ou Open XML do Office) para o Word, texto e dados tabulares para o Excel e texto para o PowerPoint e o Project. Você também pode criar manipuladores de eventos para detectar alterações na seleção do usuário. O exemplo a seguir obtém dados da seleção como texto usando o `getSelectedDataAsync` método.


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}

```

Para saber mais e obter exemplos, consulte [Ler e gravar dados na seleção ativa em um documento ou planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>Associar a uma região em um documento ou planilha

Você pode usar os `getSelectedDataAsync` métodos `setSelectedDataAsync` e para ler ou gravar na seleção *atual* do usuário em um documento, planilha ou apresentação. No entanto, para acessar a mesma região em um documento entre sessões de execução do suplemento sem exigir que o usuário faça uma seleção, primeiro você deve se associar a essa região. Você também pode se inscrever para eventos de dados e alteração de seleção para a região associada.

Você pode adicionar uma associação usando os métodos [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-), [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-) ou [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) para o objeto [Bindings](/javascript/api/office/office.bindings). Esses métodos retornam um identificador que você pode usar para acessar dados na associação ou para assinar seus eventos de alteração de dados ou alteração de seleção.

Veja a seguir um exemplo que adiciona uma associação ao texto selecionado no momento em um documento usando o `Bindings.addFromSelectionAsync` método.



```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Para saber mais e obter exemplos, consulte [Associar a regiões em um documento ou planilha](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="getting-entire-documents"></a>Obter documentos inteiros

Se o suplemento de painel de tarefas for executado no PowerPoint ou no Word, você poderá usar os métodos [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-), [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) e [File.closeAsync](/javascript/api/office/office.file#closeasync-callback-) para obter um documento ou apresentação inteira.

Ao chamar `Document.getFileAsync` você obtém uma cópia do documento em um objeto [File](/javascript/api/office/office.file) . O `File` objeto fornece acesso ao documento em "partes" representadas como objetos de [fatia](/javascript/api/office/office.slice) . Ao chamar `getFileAsync`, você pode especificar o tipo de arquivo (texto ou o formato XML do Open Office compactado) e o tamanho das fatias (até 4 MB). Para acessar o conteúdo do `File` objeto, você chama `File.getSliceAsync` o que retorna os dados brutos na propriedade [Slice. Data](/javascript/api/office/office.slice#data) . Se tiver especificado o formato compactado, você obterá os dados do arquivo como uma matriz de bytes. Se estiver transmitindo o arquivo para um serviço Web, você poderá transformar os dados brutos compactados em uma cadeia de caracteres codificada na base 64 antes do envio. Por fim, quando terminar de obter as fatias do arquivo, use `File.closeAsync` o método para fechar o documento.

Para saber mais, veja como [obter todo o documento por meio de um suplemento para PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Ler e gravar partes XML personalizadas de um documento do Word

Usando o formato de arquivo Open Office XML e controles de conteúdo, você pode adicionar partes XML personalizadas a um documento do Word e associar elementos nas partes XML a controles de conteúdo no documento. Quando você abre o documento, o Word lê e popula automaticamente os controles de conteúdo associados com dados das partes XML personalizadas. Os usuários também podem gravar dados nos controles de conteúdo. Quando o usuário salvar o documento, os dados nos controles serão salvos nas partes XML associadas. Suplementos de painel de tarefas do Word podem usar a propriedade [Document.customXmlParts](/javascript/api/office/office.document#customxmlparts) e os objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) e [CustomXmlNode](/javascript/api/office/office.customxmlnode) para ler e gravar dados dinamicamente no documento.

Partes XML personalizadas podem ser associadas a namespaces. Para obter dados de partes XML personalizadas em um namespace, use o método [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-).

Você também pode usar o método [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) para acessar partes XML personalizadas por meio de seus GUIDs. Depois de obter uma parte XML personalizada, use o método [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-) para obter os dados XML.

Para adicionar uma nova parte XML personalizada a um documento, use a `Document.customXmlParts` propriedade para obter as partes XML personalizadas que estão no documento e chame o método [CustomXmlParts. addasync](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-) .

Para saber mais sobre como trabalhar com partes XML personalizadas em um suplemento de painel de tarefas, consulte [Criar suplementos melhores para o Word com o Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Persistir configurações de suplemento


Muitas vezes, você precisa salvar dados personalizados no suplemento, como preferências do usuário ou o estado do suplemento, e acessar esses dados na próxima vez que o suplemento for aberto. Você pode usar técnicas de programação comuns para salvar os dados, como cookies do navegador ou armazenamento na Web em HTML 5. Como alternativa, se o suplemento for executado no Excel, no PowerPoint ou no Word, você poderá usar os métodos do objeto [Settings](/javascript/api/office/office.settings). Os dados criados com `Settings` o objeto são armazenados na planilha, apresentação ou documento no qual o suplemento foi inserido e salvo. Esses dados estão disponíveis apenas para o suplemento que os criou.

Para evitar viagens de ida e volta ao servidor onde o documento está armazenado, os `Settings` dados criados com o objeto são gerenciados na memória em tempo de execução. Dados de configurações salvos anteriormente são carregados na memória quando o suplemento é inicializado, e alterações nesses dados só são salvas de volta para o documento quando você chama o método [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-). Internamente, os dados são armazenados em um objeto JSON serializado como pares de nome/valor. Você usa os métodos [get](/javascript/api/office/office.settings#get-name-), [set](/javascript/api/office/office.settings#set-name--value-) e [remove](/javascript/api/office/office.settings#remove-name-) para o objeto **Settings**, para ler, gravar e excluir itens da cópia dos dados na memória. A linha de código a seguir mostra como criar uma configuração denominada `themeColor` e definir seu valor como 'verde'.




```js
Office.context.document.settings.set('themeColor', 'green');
```

Como os dados de configurações criados ou excluídos `set` com `remove` os métodos e estão agindo em uma cópia na memória dos dados, você deve chamar `saveAsync` para manter as alterações nos dados de configurações no documento com o qual o suplemento está trabalhando.

Para obter mais detalhes sobre como trabalhar com dados personalizados usando os métodos `Settings` do objeto, consulte [persistendo o estado e as configurações do suplemento](persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Ler propriedades de um documento do Project

Se o suplemento de painel de tarefas for executado no Project, o suplemento poderá ler dados de alguns dos campos de projeto, recursos e campos de tarefa do projeto ativo. Para fazer isso, você usa os métodos e os eventos do objeto [ProjectDocument](/javascript/api/office/office.document) , que estende o `Document` objeto para fornecer funcionalidade adicional específica do projeto.

Para obter exemplos de leitura de dados do Project, consulte [Criar seu primeiro suplemento de painel de tarefas do Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="permissions-model-and-governance"></a>Modelo de permissões e governança

Seu suplemento usa o `Permissions` elemento em seu manifesto para solicitar permissão para acessar o nível de funcionalidade que ele requer da API JavaScript do Office. Por exemplo, se seu suplemento requer acesso de leitura/gravação ao documento, seu manifesto deve especificar `ReadWriteDocument` como o valor de texto em seu `Permissions` elemento. Uma vez existem permissões para proteger a privacidade e a segurança do usuário, como prática recomendada, você deve solicitar o nível mínimo de permissões necessárias para seus recursos. O exemplo a seguir mostra como solicitar a permissão **ReadDocument** no manifesto de um painel de tarefas.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

Para obter mais informações, consulte [solicitar permissões para uso da API em suplementos](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## <a name="see-also"></a>Também confira

- [API JavaScript para Office](../reference/javascript-api-for-office.md)
- [Referência de esquema para os manifestos dos Suplementos do Office](../develop/add-in-manifests.md)
- [Solucionar erros de usuários com Suplementos do Office](../testing/testing-and-troubleshooting.md)
