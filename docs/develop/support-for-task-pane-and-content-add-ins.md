---
title: Suporte da API JavaScript para Office para suplementos de conteúdo e de painel de tarefas no Office 2013
description: ''
ms.date: 12/04/2017
localization_priority: Normal
ms.openlocfilehash: c9c82905a00e0cf2b2d545bb81a540e931d25407
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635913"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Suporte da API JavaScript para Office para suplementos de conteúdo e de painel de tarefas no Office 2013


Você pode usar a [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) para criar suplementos de painel de tarefas ou de conteúdo para aplicativos host do Office 2013. Os objetos e métodos que dão suporte a suplementos de conteúdo e de painel de tarefas são categorizados da seguinte forma:


1. **Objetos comuns compartilhados com outros Suplementos do Office.** Esses objetos incluem [Office](https://docs.microsoft.com/javascript/api/office), [Context](https://docs.microsoft.com/javascript/api/office/office.context) e [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult). O objeto **Office** é o objeto raiz da API JavaScript para Office. O objeto **Context** representa o ambiente de tempo de execução do suplemento. **Office** e **Context** são os objetos fundamentais para qualquer Suplemento do Office. O objeto **AsyncResult** representa os resultados de uma operação assíncrona, como os dados retornados ao método **getSelectedDataAsync**, que lê o que um usuário selecionou em um documento.
    
2.  **O objeto Document.** A maioria da API disponível para conteúdo e tarefa painel suplementos é exposta através de métodos, propriedades e eventos do objeto [Document](https://docs.microsoft.com/javascript/api/office/office.document) . Um painel de conteúdo ou a tarefa suplemento pode usar a propriedade [Office.context.document](https://docs.microsoft.com/javascript/api/office/office.context#document) para acessar o objeto de **documento** e pelo, pode acessar os membros principais da API para trabalhar com dados em documentos, como as [ligações](https://docs.microsoft.com/javascript/api/office/office.bindings) e [ CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts) objetos e os métodos [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)e [getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document#getfileasync-filetype--options--callback-) . O objeto **Document** também fornece a propriedade [mode](https://docs.microsoft.com/javascript/api/office/office.document#mode) para determinar se um documento é somente leitura ou no modo de edição, a propriedade de [url](https://docs.microsoft.com/javascript/api/office/office.document#url) para obter a URL do documento atual e para o objeto de [configurações](https://docs.microsoft.com/javascript/api/office/office.settings) de acesso. O objeto **Document** também oferece suporte adicionando manipuladores de eventos para o evento [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) , você pode detectar quando um usuário altera sua seleção no documento.
    
   Um suplemento de conteúdo ou painel de tarefas só pode acessar o objeto **Document** depois que o DOM e o ambiente de tempo de execução são carregados, normalmente no manipulador de eventos para o evento [Office.initialize](https://docs.microsoft.com/javascript/api/office). Para saber mais sobre o fluxo de eventos quando um suplemento é inicializado e como verificar se o DOM e o tempo de execução foram carregados com êxito, confira [Carregar o DOM e o ambiente de tempo de execução](loading-the-dom-and-runtime-environment.md).
    
3.  **Objetos para trabalhar com recursos específicos.** Para trabalhar com recursos específicos da API, use as seguintes objetos e métodos:
    
    - Os métodos do objeto [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings) para criar ou obter associações e os métodos e propriedades do objeto [Binding](https://docs.microsoft.com/javascript/api/office/office.binding) para trabalhar com dados.
    
    - Os objetos [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts), [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart) e objetos associados para criar e manipular partes XML personalizadas em documentos do Word.
    
    - Os objetos [File](https://docs.microsoft.com/javascript/api/office/office.file) e [Slice](https://docs.microsoft.com/javascript/api/office/office.slice) para criar uma cópia do documento inteiro, dividi-lo em partes ou "fatias" e ler ou transmitir os dados nessas fatias.
    
    - O objeto [Settings](https://docs.microsoft.com/javascript/api/office/office.settings) para salvar dados personalizados, como preferências do usuário e o estado do suplemento.
    

> [!IMPORTANT]
> Alguns membros da API não têm suporte em todos os aplicativos do Office que podem hospedar suplementos de conteúdo e de painel de tarefas. Para determinar quais membros têm suporte, confira o seguinte:

Confira um resumo do suporte à API JavaScript para Office entre os aplicativos host do Office em [Noções básicas sobre a API JavaScript para Office](understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Ler e gravar em uma seleção ativa

Você pode ler ou gravar na seleção atual do usuário em um documento, planilha ou apresentação. Dependendo do aplicativo host para o suplemento, você pode especificar o tipo de estrutura de dados para ler ou gravar como um parâmetro nos métodos [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) e [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) do objeto [Document](https://docs.microsoft.com/javascript/api/office/office.document). Por exemplo, você pode especificar qualquer tipo de dados (texto, HTML, dados tabulares ou Open XML do Office) para o Word, texto e dados tabulares para o Excel e texto para o PowerPoint e o Project. Você também pode criar manipuladores de eventos para detectar alterações na seleção do usuário. O exemplo a seguir obtém dados da seleção como texto usando o método **getSelectedDataAsync**.


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

Saiba mais e veja exemplos em [Ler e gravar dados na seleção ativa em um documento ou planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>Associar a uma região em um documento ou planilha

Você pode usar os métodos **getSelectedDataAsync** e **setSelectedDataAsync** para ler ou gravar na seleção *atual* do usuário em um documento, planilha ou apresentação. No entanto, para acessar a mesma região em um documento entre sessões de execução do suplemento sem exigir que o usuário faça uma seleção, primeiro você deve se associar a essa região. Você também pode se inscrever para eventos de dados e alteração de seleção para a região associada.

Você pode adicionar uma associação usando os métodos [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-), [addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-) ou [addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) para o objeto [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings). Esses métodos retornam um identificador que você pode usar para acessar dados na associação ou para assinar seus eventos de alteração de dados ou alteração de seleção.

A seguir está um exemplo que adiciona uma associação ao texto selecionado no momento em um documento usando o método **Bindings.addFromSelectionAsync**.



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

Saiba mais e veja exemplos em [Associar a regiões em um documento ou planilha](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="getting-entire-documents"></a>Obtendo documentos inteiros

Se o suplemento de painel de tarefas for executado no PowerPoint ou no Word, você poderá usar os métodos [Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document#getfileasync-filetype--options--callback-), [File.getSliceAsync](https://docs.microsoft.com/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) e [File.closeAsync](https://docs.microsoft.com/javascript/api/office/office.file#closeasync-callback-) para obter um documento ou apresentação inteira.

Ao chamar **Document.getFileAsync**, você obtém uma cópia do documento em um objeto [File](https://docs.microsoft.com/javascript/api/office/office.file). O objeto **File** fornece acesso ao documento em "partes" representadas como objetos [Slice](https://docs.microsoft.com/javascript/api/office/office.slice). Ao chamar **getFileAsync**, você pode especificar o tipo de arquivo (texto ou formato Open XML do Office compactado) e o tamanho das fatias (até 4 MB). Para acessar o conteúdo do objeto **File**, chame **File.getSliceAsync**, que retorna os dados brutos na propriedade [Slice.data](https://docs.microsoft.com/javascript/api/office/office.slice#data). Se tiver especificado o formato compactado, você obterá os dados do arquivo como uma matriz de bytes. Se estiver transmitindo o arquivo para um serviço Web, você poderá transformar os dados brutos compactados em uma cadeia de caracteres codificada na base 64 antes do envio. Finalmente, ao terminar de obter fatias do arquivo, use o método **File.closeAsync** para fechar o documento.

Para saber mais, veja como [obter todo o documento por meio de um suplemento para PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md). 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Lendo e gravando partes XML personalizadas de um documento do Word

Usando o formato de arquivo Open Office XML e controles de conteúdo, você pode adicionar partes XML personalizadas a um documento do Word e associar elementos nas partes XML a controles de conteúdo no documento. Quando você abre o documento, o Word lê e popula automaticamente os controles de conteúdo associados com dados das partes XML personalizadas. Os usuários também podem gravar dados nos controles de conteúdo. Quando o usuário salvar o documento, os dados nos controles serão salvos nas partes XML associadas. Suplementos de painel de tarefas do Word podem usar a propriedade [Document.customXmlParts](https://docs.microsoft.com/javascript/api/office/office.document#customxmlparts) e os objetos [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts), [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart) e [CustomXmlNode](https://docs.microsoft.com/javascript/api/office/office.customxmlnode) para ler e gravar dados dinamicamente no documento.

Partes XML personalizadas podem ser associadas a namespaces. Para obter dados de partes XML personalizadas em um namespace, use o método [CustomXmlParts.getByNamespaceAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-).

Você também pode usar o método [CustomXmlParts.getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) para acessar partes XML personalizadas por meio de seus GUIDs. Depois de obter uma parte XML personalizada, use o método [CustomXmlPart.getXmlAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-) para obter os dados XML.

Para adicionar um novo componente XML personalizado a um documento, use a propriedade **Document.customXmlParts** para bloquear as partes XML personalizadas que estão no documento e chame o método [CustomXmlParts.addAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-).

Para saber mais sobre como trabalhar com partes XML personalizadas em um suplemento de painel de tarefas, consulte [Criar suplementos melhores para o Word com o Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Persistir configurações de suplemento


Muitas vezes, você precisa salvar dados personalizados no suplemento, como preferências do usuário ou o estado do suplemento, e acessar esses dados na próxima vez que o suplemento for aberto. Você pode usar técnicas de programação comuns para salvar os dados, como cookies do navegador ou armazenamento na Web em HTML 5. Como alternativa, se o suplemento for executado no Excel, no PowerPoint ou no Word, você poderá usar os métodos do objeto [Settings](https://docs.microsoft.com/javascript/api/office/office.settings). Os dados criados com o objeto **Settings** são armazenados na planilha, na apresentação ou no documento em que o suplemento foi inserido e salvo. Esses dados estão disponíveis apenas para o suplemento que os criou.

Para evitar percursos circulares para o servidor onde o documento está armazenado, dados criados com o objeto de **configurações** são gerenciados na memória em tempo de execução. Configurações de dados são carregados na memória quando o suplemento é inicializado e altera para que dados são salvas somente volta para o documento quando você chama o método [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings#saveasync-options--callback-) foi salvo anteriormente. Internamente, os dados são armazenados em um objeto JSON serializado como pares nome/valor. Você pode usar os métodos de [obter](https://docs.microsoft.com/javascript/api/office/office.settings#get-name-), [Definir](https://docs.microsoft.com/javascript/api/office/office.settings#set-name--value-)e [Remover](https://docs.microsoft.com/javascript/api/office/office.settings#remove-name-) do objeto de **configurações** , para ler, gravar e excluir itens da cópia na memória dos dados. A linha de código a seguir mostra como criar uma configuração denominada `themeColor` e defina seu valor como 'verde'.




```js
Office.context.document.settings.set('themeColor', 'green');
```

Como os dados de configurações criados ou excluídos com os métodos **set** e **remove** atuam em uma cópia dos dados na memória, você deve chamar **saveAsync** para persistir as alterações feitas nos dados de configurações no documento com o qual o suplemento está trabalhando.

Saiba mais sobre como trabalhar com dados personalizados usando os métodos do objeto **Settings** em [Persistir o estado e as configurações do suplemento](persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Ler as propriedades de um documento de projeto

Se o suplemento de painel de tarefas for executado no Project, o suplemento poderá ler dados de alguns dos campos de projeto, recursos e campos de tarefa do projeto ativo. Para fazer isso, você usa os métodos e eventos do objeto [ProjectDocument](https://docs.microsoft.com/javascript/api/office/office.document), que estende o objeto **Document** para fornecer funcionalidade adicional específica do Project.

Para obter exemplos de leitura de dados do Project, consulte [Criar seu primeiro suplemento de painel de tarefas do Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="permissions-model-and-governance"></a>Modelo de permissões e governança

O suplemento usa o elemento **Permissions** em seu manifesto para solicitar permissão para acessar o nível de funcionalidade necessário da API JavaScript para Office. Por exemplo, se o suplemento exigir acesso de leitura/gravação ao documento, seu manifesto deverá especificar `ReadWriteDocument` como o valor de texto no elemento **Permissions**. Uma vez existem permissões para proteger a privacidade e a segurança do usuário, como prática recomendada, você deve solicitar o nível mínimo de permissões necessárias para seus recursos. O exemplo a seguir mostra como solicitar a permissão **ReadDocument** no manifesto de um painel de tarefas.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

Saiba mais em [Solicitação de permissões para uso da API em suplementos de conteúdo e de painel de tarefas](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## <a name="see-also"></a>Veja também

- [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
- [Referência de esquema para manifestos de suplementos do Office](../develop/add-in-manifests.md)
- [Solucionar erros de usuários com suplementos do Office](../testing/testing-and-troubleshooting.md)
    
