---
title: Suporte da API JavaScript para Office para suplementos de conteúdo e de painel de tarefas no Office 2013
description: Use a Office JavaScript para criar um painel de tarefas no Office 2013.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 644bc1f0759d381de412cb276a1535d2251abb0a6a0be78b45d9cc0a245758c7
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079990"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Suporte da API JavaScript para Office para suplementos de conteúdo e de painel de tarefas no Office 2013

[!include[information about the common API](../includes/alert-common-api-info.md)]

Você pode usar a [api Office JavaScript](../reference/javascript-api-for-office.md) para criar o painel de tarefas ou os complementos de conteúdo para Office aplicativos cliente 2013. Os objetos e métodos que dão suporte a suplementos de conteúdo e de painel de tarefas são categorizados da seguinte forma:

1. **Objetos comuns compartilhados com outros Office de complementos.** Esses objetos incluem [Office,](/javascript/api/office) [Contexto](/javascript/api/office/office.context)e [AsyncResult](/javascript/api/office/office.asyncresult). O `Office` objeto é o objeto raiz da API JavaScript Office JavaScript. O objeto representa o ambiente de tempo de `Context` execução do complemento. Ambos `Office` e são os objetos `Context` fundamentais para qualquer Office Add-in. O objeto representa os resultados de uma operação assíncrona, como os dados retornados ao método, que lê o que um usuário `AsyncResult` `getSelectedDataAsync` selecionou em um documento.

2. **O objeto Document.** A maioria da API disponível para o conteúdo e os complementos do painel de tarefas é exposta por meio dos métodos, propriedades e eventos do [objeto Document.](/javascript/api/office/office.document) Um add-in de conteúdo ou painel de tarefas pode usar [Office.context.docpropriedade ument](/javascript/api/office/office.context#document) doOffice.context.docpara acessar o objeto **Document** e, por meio dele, pode acessar os membros-chave da API para trabalhar com dados em documentos, como os objetos [Bindings](/javascript/api/office/office.bindings) e [CustomXmlParts,](/javascript/api/office/office.customxmlparts) e os métodos [getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_), [setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_)e [getFileAsync.](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_) O objeto também fornece a propriedade mode para determinar se um documento é somente leitura ou no modo de edição, a propriedade url para obter a URL do documento atual e o acesso ao objeto `Document` [Configurações.](/javascript/api/office/office.settings) [](/javascript/api/office/office.document#mode) [](/javascript/api/office/office.document#url) O objeto também dá suporte à adição de manipuladores de eventos para o `Document` [evento SelectionChanged,](/javascript/api/office/office.documentselectionchangedeventargs) para que você possa detectar quando um usuário altera sua seleção no documento.

   Um conteúdo ou um complemento do painel de tarefas pode acessar o objeto somente depois que o dom e o ambiente de tempo de execução foram carregados, normalmente no manipulador de eventos do `Document` [eventoOffice.initialize.](/javascript/api/office) Para saber mais sobre o fluxo de eventos quando um suplemento é inicializado e como verificar se o DOM e o tempo de execução foram carregados com êxito, confira [Carregar o DOM e o ambiente de tempo de execução](loading-the-dom-and-runtime-environment.md).

3. **Objetos para trabalhar com recursos específicos.** Para trabalhar com recursos específicos da API, use os seguintes objetos e métodos.

    - Os métodos do objeto [Bindings](/javascript/api/office/office.bindings) para criar ou obter associações e os métodos e propriedades do objeto [Binding](/javascript/api/office/office.binding) para trabalhar com dados.

    - Os objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) e objetos associados para criar e manipular partes XML personalizadas em documentos do Word.

    - Os objetos [File](/javascript/api/office/office.file) e [Slice](/javascript/api/office/office.slice) para criar uma cópia do documento inteiro, dividi-lo em partes ou "fatias" e ler ou transmitir os dados nessas fatias.

    - O objeto [Settings](/javascript/api/office/office.settings) para salvar dados personalizados, como preferências do usuário e o estado do suplemento.

> [!IMPORTANT]
> Alguns membros da API não têm suporte em todos os aplicativos do Office que podem hospedar suplementos de conteúdo e de painel de tarefas. Para determinar quais membros têm suporte, confira o seguinte:

Para um resumo do Office da API JavaScript em aplicativos Office cliente, consulte [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md).

## <a name="read-and-write-to-an-active-selection-in-a-document-spreadsheet-or-presentation"></a>Ler e gravar em uma seleção ativa em um documento, planilha ou apresentação

Você pode ler ou gravar na seleção atual do usuário em um documento, planilha ou apresentação. Dependendo do aplicativo Office do seu add-in, você pode especificar o tipo de estrutura de dados a ser lida ou escrita como um parâmetro nos métodos [getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) e [setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) do objeto [Document.](/javascript/api/office/office.document) Por exemplo, você pode especificar qualquer tipo de dados (texto, HTML, dados tabulares ou Open XML do Office) para o Word, texto e dados tabulares para o Excel e texto para o PowerPoint e o Project. Você também pode criar manipuladores de eventos para detectar alterações na seleção do usuário. O exemplo a seguir obtém dados da seleção como texto usando o `getSelectedDataAsync` método.


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

## <a name="bind-to-a-region-in-a-document-or-spreadsheet"></a>Vincular a uma região em um documento ou planilha

Você pode usar os métodos e para ler ou gravar na seleção atual do usuário em `getSelectedDataAsync` `setSelectedDataAsync` um documento, planilha ou apresentação.  No entanto, para acessar a mesma região em um documento entre sessões de execução do suplemento sem exigir que o usuário faça uma seleção, primeiro você deve se associar a essa região. Você também pode se inscrever para eventos de dados e alteração de seleção para a região associada.

Você pode adicionar uma associação usando os métodos [addFromNamedItemAsync](/javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_), [addFromPromptAsync](/javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_) ou [addFromSelectionAsync](/javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_) para o objeto [Bindings](/javascript/api/office/office.bindings). Esses métodos retornam um identificador que você pode usar para acessar dados na associação ou para assinar seus eventos de alteração de dados ou alteração de seleção.

A seguir, um exemplo que adiciona uma associação ao texto selecionado no momento em um documento, usando o `Bindings.addFromSelectionAsync` método.

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

Se o suplemento de painel de tarefas for executado no PowerPoint ou no Word, você poderá usar os métodos [Document.getFileAsync](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_), [File.getSliceAsync](/javascript/api/office/office.file#getSliceAsync_sliceIndex__callback_) e [File.closeAsync](/javascript/api/office/office.file#closeAsync_callback_) para obter um documento ou apresentação inteira.

Ao `Document.getFileAsync` chamar, você obterá uma cópia do documento em um [objeto File.](/javascript/api/office/office.file) O `File` objeto fornece acesso ao documento em "partes" representadas como objetos [Slice.](/javascript/api/office/office.slice) Ao chamar , você pode especificar o tipo de arquivo (texto ou formato XML compactado Office) e o tamanho das `getFileAsync` fatias (até 4 MB). Para acessar o conteúdo do objeto, você chama o que retorna os dados `File` `File.getSliceAsync` brutos na [propriedade Slice.data.](/javascript/api/office/office.slice#data) Se tiver especificado o formato compactado, você obterá os dados do arquivo como uma matriz de bytes. Se estiver transmitindo o arquivo para um serviço Web, você poderá transformar os dados brutos compactados em uma cadeia de caracteres codificada na base 64 antes do envio. Por fim, quando terminar de obter fatias do arquivo, use o `File.closeAsync` método para fechar o documento.

Para saber mais, veja como [obter todo o documento por meio de um suplemento para PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).

## <a name="read-and-write-custom-xml-parts-of-a-word-document"></a>Ler e gravar partes XML personalizadas de um documento do Word

Usando o formato de arquivo Open Office XML e controles de conteúdo, você pode adicionar partes XML personalizadas a um documento do Word e associar elementos nas partes XML a controles de conteúdo no documento. Quando você abre o documento, o Word lê e popula automaticamente os controles de conteúdo associados com dados das partes XML personalizadas. Os usuários também podem gravar dados nos controles de conteúdo. Quando o usuário salvar o documento, os dados nos controles serão salvos nas partes XML associadas. Suplementos de painel de tarefas do Word podem usar a propriedade [Document.customXmlParts](/javascript/api/office/office.document#customXmlParts) e os objetos [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) e [CustomXmlNode](/javascript/api/office/office.customxmlnode) para ler e gravar dados dinamicamente no documento.

Partes XML personalizadas podem ser associadas a namespaces. Para obter dados de partes XML personalizadas em um namespace, use o método [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getByNamespaceAsync_ns__options__callback_).

Você também pode usar o método [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getByIdAsync_id__options__callback_) para acessar partes XML personalizadas por meio de seus GUIDs. Depois de obter uma parte XML personalizada, use o método [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getXmlAsync_options__callback_) para obter os dados XML.

Para adicionar uma nova parte XML personalizada a um documento, use a propriedade para obter as partes XML personalizadas que estão no documento e chame o `Document.customXmlParts` [método CustomXmlParts.addAsync.](/javascript/api/office/office.customxmlparts#addAsync_xml__options__callback_)

Para saber mais sobre como trabalhar com partes XML personalizadas em um suplemento de painel de tarefas, consulte [Criar suplementos melhores para o Word com o Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).

## <a name="persisting-add-in-settings"></a>Persistir configurações de suplemento

Muitas vezes, você precisa salvar dados personalizados no suplemento, como preferências do usuário ou o estado do suplemento, e acessar esses dados na próxima vez que o suplemento for aberto. Você pode usar técnicas de programação comuns para salvar os dados, como cookies do navegador ou armazenamento na Web em HTML 5. Como alternativa, se o suplemento for executado no Excel, no PowerPoint ou no Word, você poderá usar os métodos do objeto [Settings](/javascript/api/office/office.settings). Os dados criados com o objeto são armazenados na planilha, apresentação ou documento em que o complemento foi `Settings` inserido e salvo. Esses dados estão disponíveis apenas para o suplemento que os criou.

Para evitar idas e voltas para o servidor onde o documento está armazenado, os dados criados com o objeto são gerenciados na `Settings` memória em tempo de executar. Dados de configurações salvos anteriormente são carregados na memória quando o suplemento é inicializado, e alterações nesses dados só são salvas de volta para o documento quando você chama o método [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_). Internamente, os dados são armazenados em um objeto JSON serializado como pares de nome/valor. Você usa os métodos [get](/javascript/api/office/office.settings#get_name_), [set](/javascript/api/office/office.settings#set_name__value_) e [remove](/javascript/api/office/office.settings#remove_name_) para o objeto **Settings**, para ler, gravar e excluir itens da cópia dos dados na memória. A linha de código a seguir mostra como criar uma configuração denominada `themeColor` e definir seu valor como 'verde'.

```js
Office.context.document.settings.set('themeColor', 'green');
```

Como os dados de configurações criados ou excluídos com os métodos e estão agindo em uma cópia na memória dos dados, você deve chamar para persistir alterações nas configurações de dados no documento com o que o seu complemento está `set` `remove` `saveAsync` trabalhando.

Para obter mais detalhes sobre como trabalhar com dados personalizados usando os métodos do objeto, consulte `Settings` [Persisting add-in state and settings](persisting-add-in-state-and-settings.md).

## <a name="read-properties-of-a-project-document"></a>Ler propriedades de um documento de projeto

Se o suplemento de painel de tarefas for executado no Project, o suplemento poderá ler dados de alguns dos campos de projeto, recursos e campos de tarefa do projeto ativo. Para fazer isso, use os métodos e eventos do [objeto ProjectDocument,](/javascript/api/office/office.document) que estende o objeto para fornecer funcionalidades Project `Document` específicas.

Para obter exemplos de leitura de dados do Project, consulte [Criar seu primeiro suplemento de painel de tarefas do Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).

## <a name="permissions-model-and-governance"></a>Modelo de permissões e governança

Seu add-in usa o elemento em seu manifesto para solicitar permissão para acessar o nível de funcionalidade que ele exige da API `Permissions` Office JavaScript. Por exemplo, se o seu complemento exigir acesso de leitura/gravação ao documento, seu manifesto deverá especificar como o valor `ReadWriteDocument` de texto em seu `Permissions` elemento. Uma vez existem permissões para proteger a privacidade e a segurança do usuário, como prática recomendada, você deve solicitar o nível mínimo de permissões necessárias para seus recursos. O exemplo a seguir mostra como solicitar a permissão **ReadDocument** no manifesto de um painel de tarefas.

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

Para obter mais informações, consulte [Solicitando permissões para uso da API em complementos](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).

## <a name="see-also"></a>Confira também

- [API JavaScript para Office](../reference/javascript-api-for-office.md)
- [Referência de esquema para os manifestos dos Suplementos do Office](../develop/add-in-manifests.md)
- [Solucionar erros de usuários com Suplementos do Office](../testing/testing-and-troubleshooting.md)
