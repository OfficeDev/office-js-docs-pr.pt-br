---
title: Suporte da API JavaScript para Office para suplementos de conteúdo e de painel de tarefas no Office 2013
description: Use a API JavaScript do Office para criar um painel de tarefas no Office 2013.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: a6072538fe7328a71767394adf67398ebe4f0911
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422863"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Suporte da API JavaScript para Office para suplementos de conteúdo e de painel de tarefas no Office 2013

[!include[information about the common API](../includes/alert-common-api-info.md)]

Você pode usar a [API JavaScript do Office](../reference/javascript-api-for-office.md) para criar suplementos de conteúdo ou painel de tarefas para aplicativos cliente do Office 2013. Os objetos e métodos que dão suporte a suplementos de conteúdo e de painel de tarefas são categorizados da seguinte forma:

1. **Objetos comuns compartilhados com outros Suplementos do Office.** Esses objetos [incluem Office](/javascript/api/office), [Context](/javascript/api/office/office.context) e [AsyncResult](/javascript/api/office/office.asyncresult). O `Office` objeto é o objeto raiz da API JavaScript do Office. O `Context` objeto representa o ambiente de runtime do suplemento. Ambos `Office` e `Context` são os objetos fundamentais para qualquer Suplemento do Office. O `AsyncResult` objeto representa os resultados de uma operação assíncrona, `getSelectedDataAsync` como os dados retornados ao método, que lê o que um usuário selecionou em um documento.

2. **O objeto Document.** A maioria da API disponível para suplementos de conteúdo e painel de tarefas é exposta por meio dos métodos, propriedades e eventos do [objeto Document](/javascript/api/office/office.document) . Um suplemento de conteúdo ou painel de tarefas pode usar a propriedade [Office.context.document](/javascript/api/office/office.context#office-office-context-document-member) para acessar o objeto **Document** e, por meio dele, pode acessar os principais membros da API para trabalhar com dados em documentos, como os objetos [Bindings](/javascript/api/office/office.bindings) e [CustomXmlParts](/javascript/api/office/office.customxmlparts) , e os métodos [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)), [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) e [getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) . O `Document` objeto também fornece a [](/javascript/api/office/office.document#office-office-document-mode-member) propriedade de modo para determinar se um documento é somente leitura ou no modo de edição, a propriedade [de URL](/javascript/api/office/office.document#office-office-document-url-member) para obter a URL do documento atual e o acesso ao objeto [Settings](/javascript/api/office/office.settings). O `Document` objeto também dá suporte à adição de manipuladores de eventos para o [evento SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) , para que você possa detectar quando um usuário altera sua seleção no documento.

   Um suplemento de `Document` conteúdo ou painel de tarefas pode acessar o objeto somente depois que o DOM e o ambiente de runtime tiverem sido carregados, normalmente no manipulador de eventos do evento [Office.initialize](/javascript/api/office) . Para saber mais sobre o fluxo de eventos quando um suplemento é inicializado e como verificar se o DOM e o tempo de execução foram carregados com êxito, confira [Carregar o DOM e o ambiente de tempo de execução](loading-the-dom-and-runtime-environment.md).

3. **Objetos para trabalhar com recursos específicos.** Para trabalhar com recursos específicos da API, use os seguintes objetos e métodos.

    - Os métodos do objeto [Bindings](/javascript/api/office/office.bindings) para criar ou obter associações e os métodos e propriedades do objeto [Binding](/javascript/api/office/office.binding) para trabalhar com dados.

    - Os objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) e objetos associados para criar e manipular partes XML personalizadas em documentos do Word.

    - Os objetos [File](/javascript/api/office/office.file) e [Slice](/javascript/api/office/office.slice) para criar uma cópia do documento inteiro, dividi-lo em partes ou "fatias" e ler ou transmitir os dados nessas fatias.

    - O objeto [Settings](/javascript/api/office/office.settings) para salvar dados personalizados, como preferências do usuário e o estado do suplemento.

> [!IMPORTANT]
> Alguns membros da API não têm suporte em todos os aplicativos do Office que podem hospedar suplementos de conteúdo e de painel de tarefas. Para determinar quais membros têm suporte, confira o seguinte:

Para obter um resumo do suporte à API JavaScript do Office em aplicativos cliente do Office, consulte [Noções básicas sobre a API JavaScript do Office](understanding-the-javascript-api-for-office.md).

## <a name="read-and-write-to-an-active-selection-in-a-document-spreadsheet-or-presentation"></a>Ler e gravar em uma seleção ativa em um documento, planilha ou apresentação

Você pode ler ou gravar na seleção atual do usuário em um documento, planilha ou apresentação. Dependendo do aplicativo do Office para seu suplemento, você pode especificar o tipo de estrutura de dados a ser lida ou gravada como um parâmetro nos métodos [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) e [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) do objeto [Document](/javascript/api/office/office.document) . Por exemplo, você pode especificar qualquer tipo de dados (texto, HTML, dados tabulares ou Open XML do Office) para o Word, texto e dados tabulares para o Excel e texto para o PowerPoint e o Project. Você também pode criar manipuladores de eventos para detectar alterações na seleção do usuário. O exemplo a seguir obtém dados da seleção como texto usando o `getSelectedDataAsync` método.


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

## <a name="bind-to-a-region-in-a-document-or-spreadsheet"></a>Associar a uma região em um documento ou planilha

Você pode usar os `getSelectedDataAsync` métodos `setSelectedDataAsync` e ler ou gravar na seleção atual do usuário em um  documento, planilha ou apresentação. No entanto, para acessar a mesma região em um documento entre sessões de execução do suplemento sem exigir que o usuário faça uma seleção, primeiro você deve se associar a essa região. Você também pode se inscrever para eventos de dados e alteração de seleção para a região associada.

Você pode adicionar uma associação usando os métodos [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)), [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)) ou [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) para o objeto [Bindings](/javascript/api/office/office.bindings). Esses métodos retornam um identificador que você pode usar para acessar dados na associação ou para assinar seus eventos de alteração de dados ou alteração de seleção.

A seguir está um exemplo que adiciona uma associação ao texto selecionado no momento em um documento, usando o `Bindings.addFromSelectionAsync` método.

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

## <a name="get-entire-documents"></a>Obter documentos inteiros

Se o suplemento de painel de tarefas for executado no PowerPoint ou no Word, você poderá usar os métodos [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)), [File.getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) e [File.closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)) para obter um documento ou apresentação inteira.

Ao chamar, `Document.getFileAsync` você obtém uma cópia do documento em um [objeto](/javascript/api/office/office.file) File. O `File` objeto fornece acesso ao documento em "partes" representadas como [objetos Slice](/javascript/api/office/office.slice) . Ao chamar `getFileAsync`, você pode especificar o tipo de arquivo (texto ou formato OPEN XML do Office compactado) e o tamanho das fatias (até 4 MB). Para acessar o conteúdo do objeto `File` , `File.getSliceAsync` chame o que retorna os dados brutos na [propriedade Slice.data](/javascript/api/office/office.slice#office-office-slice-data-member) . Se tiver especificado o formato compactado, você obterá os dados do arquivo como uma matriz de bytes. Se estiver transmitindo o arquivo para um serviço Web, você poderá transformar os dados brutos compactados em uma cadeia de caracteres codificada na base 64 antes do envio. Por fim, quando terminar de obter fatias do arquivo, use o `File.closeAsync` método para fechar o documento.

Para saber mais, veja como [obter todo o documento por meio de um suplemento para PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).

## <a name="read-and-write-custom-xml-parts-of-a-word-document"></a>Ler e gravar partes XML personalizadas de um documento do Word

Usando o formato de arquivo Open Office XML e controles de conteúdo, você pode adicionar partes XML personalizadas a um documento do Word e associar elementos nas partes XML a controles de conteúdo no documento. Quando você abre o documento, o Word lê e popula automaticamente os controles de conteúdo associados com dados das partes XML personalizadas. Os usuários também podem gravar dados nos controles de conteúdo. Quando o usuário salvar o documento, os dados nos controles serão salvos nas partes XML associadas. Suplementos de painel de tarefas do Word podem usar a propriedade [Document.customXmlParts](/javascript/api/office/office.document#office-office-document-customxmlparts-member) e os objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) e [CustomXmlNode](/javascript/api/office/office.customxmlnode) para ler e gravar dados dinamicamente no documento.

Partes XML personalizadas podem ser associadas a namespaces. Para obter dados de partes XML personalizadas em um namespace, use o método [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbynamespaceasync-member(1)).

Você também pode usar o método [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)) para acessar partes XML personalizadas por meio de seus GUIDs. Depois de obter uma parte XML personalizada, use o método [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#office-office-customxmlpart-getxmlasync-member(1)) para obter os dados XML.

Para adicionar uma nova parte XML personalizada a um documento, use `Document.customXmlParts` a propriedade para obter as partes XML personalizadas que estão no documento e chame o [método CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-addasync-member(1)) .

Para obter informações detalhadas sobre como gerenciar partes XML personalizadas com um suplemento do painel de tarefas, consulte Entender quando e como usar o [Office Open XML em seu suplemento do Word](../word/create-better-add-ins-for-word-with-office-open-xml.md).

## <a name="persisting-add-in-settings"></a>Persistir configurações de suplemento

Muitas vezes, você precisa salvar dados personalizados no suplemento, como preferências do usuário ou o estado do suplemento, e acessar esses dados na próxima vez que o suplemento for aberto. Você pode usar técnicas de programação comuns para salvar os dados, como cookies do navegador ou armazenamento na Web em HTML 5. Como alternativa, se o suplemento for executado no Excel, no PowerPoint ou no Word, você poderá usar os métodos do objeto [Settings](/javascript/api/office/office.settings). Os dados criados com `Settings` o objeto são armazenados na planilha, apresentação ou documento no qual o suplemento foi inserido e salvo. Esses dados estão disponíveis apenas para o suplemento que os criou.

Para evitar idas e voltas para o servidor em que o documento está armazenado, `Settings` os dados criados com o objeto são gerenciados na memória em tempo de execução. Dados de configurações salvos anteriormente são carregados na memória quando o suplemento é inicializado, e alterações nesses dados só são salvas de volta para o documento quando você chama o método [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)). Internamente, os dados são armazenados em um objeto JSON serializado como pares de nome/valor. Você usa os métodos [get](/javascript/api/office/office.settings#office-office-settings-get-member(1)), [set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) e [remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) para o objeto **Settings**, para ler, gravar e excluir itens da cópia dos dados na memória. A linha de código a seguir mostra como criar uma configuração denominada `themeColor` e definir seu valor como 'verde'.

```js
Office.context.document.settings.set('themeColor', 'green');
```

`set` `remove` Como os dados de configurações criados ou excluídos com os métodos estão agindo em uma cópia na memória dos dados, `saveAsync` você deve chamar para persistir as alterações nos dados de configurações no documento com o qual o suplemento está trabalhando.

Para obter mais detalhes sobre como trabalhar com dados personalizados usando os métodos `Settings` do objeto, consulte [Persistir o estado e as configurações do suplemento](persisting-add-in-state-and-settings.md).

## <a name="read-properties-of-a-project-document"></a>Ler propriedades de um documento de projeto

Se o suplemento de painel de tarefas for executado no Project, o suplemento poderá ler dados de alguns dos campos de projeto, recursos e campos de tarefa do projeto ativo. Para fazer isso, use os métodos e eventos do objeto [ProjectDocument](/javascript/api/office/office.document) , `Document` que estende o objeto para fornecer funcionalidade adicional específica do Project.

Para obter exemplos de leitura de dados do Project, consulte [Criar seu primeiro suplemento de painel de tarefas do Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).

## <a name="permissions-model-and-governance"></a>Modelo de permissões e governança

Seu suplemento usa o elemento em `Permissions` seu manifesto para solicitar permissão para acessar o nível de funcionalidade que ele requer da API JavaScript do Office. Por exemplo, se o suplemento exigir acesso de leitura/gravação ao documento, `ReadWriteDocument` seu manifesto deverá especificar como o valor de texto em seu `Permissions` elemento. Uma vez existem permissões para proteger a privacidade e a segurança do usuário, como prática recomendada, você deve solicitar o nível mínimo de permissões necessárias para seus recursos. O exemplo a seguir mostra como solicitar a permissão **ReadDocument** no manifesto de um painel de tarefas.

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

Para obter mais informações, consulte [Solicitando permissões para uso de API em suplementos](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).

## <a name="see-also"></a>Confira também

- [API JavaScript para Office](../reference/javascript-api-for-office.md)
- [Referência de esquema para os manifestos dos Suplementos do Office](../develop/add-in-manifests.md)
- [Solucionar erros de usuários com suplementos do Office](../testing/testing-and-troubleshooting.md)
- [Runtimes em Suplementos do Office](../testing/runtimes.md)
