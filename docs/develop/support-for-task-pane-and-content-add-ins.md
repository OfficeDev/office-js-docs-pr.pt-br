---
title: Suporte da API JavaScript para Office para suplementos de conte?do e de painel de tarefas no Office 2013
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 2aab577e3536ed11c8f2e9810f6f200bdf5d1768
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Suporte da API JavaScript para Office para suplementos de conte?do e de painel de tarefas no Office 2013


Voc? pode usar a [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) para criar suplementos de painel de tarefas ou de conte?do para aplicativos host do Office 2013. Os objetos e m?todos que d?o suporte a suplementos de conte?do e de painel de tarefas s?o categorizados da seguinte forma:


1. **Objetos comuns compartilhados com outros Suplementos do Office.** Esses objetos incluem [Office](https://dev.office.com/reference/add-ins/shared/office), [Context](https://dev.office.com/reference/add-ins/shared/office.context) e [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult). O objeto **Office** ? o objeto raiz da API JavaScript para Office. O objeto **Context** representa o ambiente de tempo de execu??o do suplemento. **Office** e **Context** s?o os objetos fundamentais para qualquer Suplemento do Office. O objeto **AsyncResult** representa os resultados de uma opera??o ass?ncrona, como os dados retornados ao m?todo **getSelectedDataAsync**, que l? o que um usu?rio selecionou em um documento.
    
2.  **O objeto Documento.** A maioria das APIs dispon?veis para suplementos do painel de tarefas e conte?do s?o expostas por meio dos m?todos, propriedades e eventos do objeto [Documento](https://dev.office.com/reference/add-ins/shared/document). Um suplemento do painel de tarefas ou conte?do pode usar a propriedade [Office.context.document](https://dev.office.com/reference/add-ins/shared/office.context.document) para acessar o objeto **Documento** e, atrav?s dele, ? poss?vel acessar os membros-chave da API para trabalhar com dados em documentos, como os objetos [Liga??es](https://dev.office.com/reference/add-ins/shared/bindings.bindings) e [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts), e os m?todos [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) e [getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync). O objeto **Documento** tamb?m fornece a propriedade [modo](https://dev.office.com/reference/add-ins/shared/document.mode) para determinar se um documento ? somente leitura ou se est? em modo de edi??o, a propriedade [url](https://dev.office.com/reference/add-ins/shared/document.url) para obter o URL do documento atual e acesso ao objeto [Configura??es](https://dev.office.com/reference/add-ins/shared/settings). O objeto **Documento** tamb?m suporta a adi??o de manipuladores de eventos para o evento [SelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) para que voc? possa detectar quando um usu?rio altera a sele??o no documento.
    
   Um suplemento de conte?do ou painel de tarefas s? pode acessar o objeto **Document** depois que o DOM e o ambiente de tempo de execu??o s?o carregados, normalmente no manipulador de eventos para o evento [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize). Para saber mais sobre o fluxo de eventos quando um suplemento ? inicializado e como verificar se o DOM e o tempo de execu??o foram carregados com ?xito, confira [Carregar o DOM e o ambiente de tempo de execu??o](loading-the-dom-and-runtime-environment.md).
    
3.  **Objetos para trabalhar com recursos espec?ficos.** Para trabalhar com recursos espec?ficos da API, use as seguintes objetos e m?todos:
    
    - Os m?todos do objeto [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) para criar ou obter associa??es e os m?todos e propriedades do objeto [Binding](https://dev.office.com/reference/add-ins/shared/binding) para trabalhar com dados.
    
    - Os objetos [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) e objetos associados para criar e manipular partes XML personalizadas em documentos do Word.
    
    - Os objetos [File](https://dev.office.com/reference/add-ins/shared/file) e [Slice](https://dev.office.com/reference/add-ins/shared/slice) para criar uma c?pia do documento inteiro, dividi-lo em partes ou "fatias" e ler ou transmitir os dados nessas fatias.
    
    - O objeto [Settings](https://dev.office.com/reference/add-ins/shared/settings) para salvar dados personalizados, como prefer?ncias do usu?rio e o estado do suplemento.
    

> [!IMPORTANT]
> Alguns membros da API n?o t?m suporte em todos os aplicativos do Office que podem hospedar suplementos de conte?do e de painel de tarefas. Para determinar quais membros t?m suporte, confira o seguinte:

Confira um resumo do suporte ? API JavaScript para Office entre os aplicativos host do Office em [No??es b?sicas sobre a API JavaScript para Office](understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Ler e gravar em uma sele??o ativa

Voc? pode ler ou gravar na sele??o atual do usu?rio em um documento, planilha ou apresenta??o. Dependendo do aplicativo host para o suplemento, voc? pode especificar o tipo de estrutura de dados para ler ou gravar como um par?metro nos m?todos [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) e [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) do objeto [Document](https://dev.office.com/reference/add-ins/shared/document). Por exemplo, voc? pode especificar qualquer tipo de dados (texto, HTML, dados tabulares ou Open XML do Office) para o Word, texto e dados tabulares para o Excel e texto para o PowerPoint e o Project. Voc? tamb?m pode criar manipuladores de eventos para detectar altera??es na sele??o do usu?rio. O exemplo a seguir obt?m dados da sele??o como texto usando o m?todo **getSelectedDataAsync**.


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

Saiba mais e veja exemplos em [Ler e gravar dados na sele??o ativa em um documento ou planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>Associar a uma regi?o em um documento ou planilha

Voc? pode usar os m?todos **getSelectedDataAsync** e **setSelectedDataAsync** para ler ou gravar na sele??o *atual* do usu?rio em um documento, planilha ou apresenta??o. No entanto, para acessar a mesma regi?o em um documento entre sess?es de execu??o do suplemento sem exigir que o usu?rio fa?a uma sele??o, primeiro voc? deve se associar a essa regi?o. Voc? tamb?m pode se inscrever para eventos de dados e altera??o de sele??o para a regi?o associada.

Voc? pode adicionar uma associa??o usando os m?todos [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync), [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync) ou [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) para o objeto [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings). Esses m?todos retornam um identificador que voc? pode usar para acessar dados na associa??o ou para assinar seus eventos de altera??o de dados ou altera??o de sele??o.

A seguir est? um exemplo que adiciona uma associa??o ao texto selecionado no momento em um documento usando o m?todo **Bindings.addFromSelectionAsync**.



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

Saiba mais e veja exemplos em [Associar a regi?es em um documento ou planilha](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="getting-entire-documents"></a>Obter documentos inteiros

Se o suplemento de painel de tarefas for executado no PowerPoint ou no Word, voc? poder? usar os m?todos [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync), [File.getSliceAsync](https://dev.office.com/reference/add-ins/shared/file.getsliceasync) e [File.closeAsync](https://dev.office.com/reference/add-ins/shared/file.closeasync) para obter um documento ou apresenta??o inteira.

Ao chamar **Document.getFileAsync**, voc? obt?m uma c?pia do documento em um objeto [File](https://dev.office.com/reference/add-ins/shared/file). O objeto **File** fornece acesso ao documento em "partes" representadas como objetos [Slice](https://dev.office.com/reference/add-ins/shared/document). Ao chamar **getFileAsync**, voc? pode especificar o tipo de arquivo (texto ou formato Open XML do Office compactado) e o tamanho das fatias (at? 4 MB). Para acessar o conte?do do objeto **File**, chame **File.getSliceAsync**, que retorna os dados brutos na propriedade [Slice.data](https://dev.office.com/reference/add-ins/shared/slice.data). Se tiver especificado o formato compactado, voc? obter? os dados do arquivo como uma matriz de bytes. Se estiver transmitindo o arquivo para um servi?o Web, voc? poder? transformar os dados brutos compactados em uma cadeia de caracteres codificada na base 64 antes do envio. Finalmente, ao terminar de obter fatias do arquivo, use o m?todo **File.closeAsync** para fechar o documento.

Para saber mais, veja como [Obter todo o documento por meio de um suplemento para PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md). 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Ler e gravar partes XML personalizadas de um documento do Word

Usando o formato de arquivo Open Office XML e controles de conte?do, voc? pode adicionar partes XML personalizadas a um documento do Word e associar elementos nas partes XML a controles de conte?do no documento. Quando voc? abre o documento, o Word l? e popula automaticamente os controles de conte?do associados com dados das partes XML personalizadas. Os usu?rios tamb?m podem gravar dados nos controles de conte?do. Quando o usu?rio salvar o documento, os dados nos controles ser?o salvos nas partes XML associadas. Suplementos de painel de tarefas do Word podem usar a propriedade [Document.customXmlParts](https://dev.office.com/reference/add-ins/shared/document.customxmlparts) e os objetos [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) e [CustomXmlNode](https://dev.office.com/reference/add-ins/shared/customxmlnode.customxmlnode) para ler e gravar dados dinamicamente no documento.

Partes XML personalizadas podem ser associadas a namespaces. Para obter dados de partes XML personalizadas em um namespace, use o m?todo [CustomXmlParts.getByNamespaceAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbynamespaceasync).

Voc? tamb?m pode usar o m?todo [CustomXmlParts.getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync) para acessar partes XML personalizadas por meio de seus GUIDs. Depois de obter uma parte XML personalizada, use o m?todo [CustomXmlPart.getXmlAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.getxmlasync) para obter os dados XML.

Para adicionar um novo componente XML personalizado a um documento, use a propriedade **Document.customXmlParts** para bloquear as partes XML personalizadas que est?o no documento e chame o m?todo [CustomXmlParts.addAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.addasync).

Para saber mais sobre como trabalhar com partes XML personalizadas em um suplemento de painel de tarefas, consulte [Criar suplementos melhores para o Word com o Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Persistir configura??es de suplemento


Muitas vezes, voc? precisa salvar dados personalizados no suplemento, como prefer?ncias do usu?rio ou o estado do suplemento, e acessar esses dados na pr?xima vez que o suplemento for aberto. Voc? pode usar t?cnicas de programa??o comuns para salvar os dados, como cookies do navegador ou armazenamento na Web em HTML 5. Como alternativa, se o suplemento for executado no Excel, no PowerPoint ou no Word, voc? poder? usar os m?todos do objeto [Settings](https://dev.office.com/reference/add-ins/shared/settings). Os dados criados com o objeto **Settings** s?o armazenados na planilha, na apresenta??o ou no documento em que o suplemento foi inserido e salvo. Esses dados est?o dispon?veis apenas para o suplemento que os criou.

Para evitar idas e voltas ao servidor em que o documento est? armazenado, dados criados com o objeto **Settings** s?o gerenciados na mem?ria em tempo de execu??o. Dados de configura??es salvos anteriormente s?o carregados na mem?ria quando o suplemento ? inicializado, e altera??es nesses dados s? s?o salvas de volta para o documento quando voc? chama o m?todo [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync). Internamente, os dados s?o armazenados em um objeto JSON serializado como pares de nome/valor. Voc? usa os m?todos [get](https://dev.office.com/reference/add-ins/shared/settings.get), [set](https://dev.office.com/reference/add-ins/shared/settings.set) e [remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) para o objeto **Settings**, para ler, gravar e excluir itens da c?pia dos dados na mem?ria. A linha de c?digo a seguir mostra como criar uma configura??o denominada `themeColor` e definir seu valor como 'verde'.




```js
Office.context.document.settings.set('themeColor', 'green');
```

Como os dados de configura??es criados ou exclu?dos com os m?todos **set** e **remove** atuam em uma c?pia dos dados na mem?ria, voc? deve chamar **saveAsync** para persistir as altera??es feitas nos dados de configura??es no documento com o qual o suplemento est? trabalhando.

Saiba mais sobre como trabalhar com dados personalizados usando os m?todos do objeto **Settings** em [Persistir o estado e as configura??es do suplemento](persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Ler propriedades de um documento do Project

Se o suplemento de painel de tarefas for executado no Project, o suplemento poder? ler dados de alguns dos campos de projeto, recursos e campos de tarefa do projeto ativo. Para fazer isso, voc? usa os m?todos e eventos do objeto [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument), que estende o objeto **Document** para fornecer funcionalidade adicional espec?fica do Project.

Veja exemplos de leitura de dados do Project em [Criar seu primeiro suplemento de painel de tarefas do Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="permissions-model-and-governance"></a>Modelo de permiss?es e governan?a

O suplemento usa o elemento **Permissions** em seu manifesto para solicitar permiss?o para acessar o n?vel de funcionalidade necess?rio da API JavaScript para Office. Por exemplo, se o suplemento exigir acesso de leitura/grava??o ao documento, seu manifesto dever? especificar `ReadWriteDocument` como o valor de texto no elemento **Permissions**. Uma vez existem permiss?es para proteger a privacidade e a seguran?a do usu?rio, como pr?tica recomendada, voc? deve solicitar o n?vel m?nimo de permiss?es necess?rias para seus recursos. O exemplo a seguir mostra como solicitar a permiss?o **ReadDocument** no manifesto de um painel de tarefas.


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

Saiba mais em [Solicita??o de permiss?es para uso da API em suplementos de conte?do e de painel de tarefas](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## <a name="see-also"></a>Veja tamb?m

- [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)
- [Refer?ncia de esquema para manifestos de suplementos do Office](../develop/add-in-manifests.md)
- [Solucionar erros de usu?rios com suplementos do Office](../testing/testing-and-troubleshooting.md)
    
