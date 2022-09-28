---
title: Visão geral da programação da API JavaScript do OneNote
description: Saiba mais sobre a API JavaScript do OneNote para suplementos do OneNote na Web.
ms.date: 07/18/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: d44a01cf0f676057ca072cff74e2e80057f645f4
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092907"
---
# <a name="onenote-javascript-api-programming-overview"></a>Visão geral da programação da API JavaScript do OneNote

OneNote introduces a JavaScript API for OneNote add-ins on the web. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a>Componentes de um suplemento do Office

Os suplementos consistem de dois componentes básicos:

- A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote on the web, the web application displays in a browser control or iframe.

- An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.

### <a name="office-add-in--manifest--webpage"></a>Suplemento do Office = Manifesto + Página da Web

![Um suplemento do Office consiste em um manifesto e uma página da Web.](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>Usar a API JavaScript

Add-ins use the runtime context of the Office application to access the JavaScript API. The API has two layers:

- Uma **API específica do aplicativo** para operações específicas do OneNote, acessada por meio do objeto `Application`.
- Uma **API comum** compartilhada entre aplicativos do Office, acessada por meio do objeto `Document`.

### <a name="accessing-the-application-specific-api-through-the-application-object"></a>Acessar uma API específica do aplicativo por meio do objeto *Aplicativo*.

Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With application-specific APIs, you run batch operations on proxy objects. The basic flow goes something like this:

1. Obtenha a instância do aplicativo do contexto.

2. Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.

3. Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.

   > [!NOTE]
   > Chamadas de método para a API (como `context.application.getActiveSection().pages;`) também são adicionadas à fila.

4. Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.

Por exemplo:

```js
async function getPagesInSection() {
    await OneNote.run(async (context) => {

        // Get the pages in the current section.
        const pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        await context.sync();
            
        // Read the id and title of each page.
        $.each(pages.items, function(index, page) {
            let pageId = page.id;
            let pageTitle = page.title;
            console.log(pageTitle + ': ' + pageId);
        });
    });
}
```

Confira [Usando o modelo de API específica do aplicativo](../develop/application-specific-api-model.md) para saber mais sobre o padrão `load`/`sync` e outras práticas comuns nas APIs de JavaScript do OneNote.

Você pode encontrar objetos do OneNote e operações compatíveis na [Referência API](../reference/overview/onenote-add-ins-javascript-reference.md).

#### <a name="onenote-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do OneNote

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets).

### <a name="accessing-the-common-api-through-the-document-object"></a>Acessar a API comum por meio do objeto *Documento*

Use o objeto `Document`para acessar a API comum, como os métodos [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) e [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)).

Por exemplo:  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            const error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

Os suplementos do OneNote dão suporte apenas às APIs comuns a seguir.

| API | Observações |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) | Somente `Office.CoercionType.Text` e `Office.CoercionType.Matrix` |
| [Office.context.document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) | Somente `Office.CoercionType.Text`, `Office.CoercionType.Image` e `Office.CoercionType.Html` |
| [const mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#office-office-settings-get-member(1)) | As configurações são compatíveis apenas com os suplementos de conteúdo |
| [Office.context.document.settings.set(nome, valor);](/javascript/api/office/office.settings#office-office-settings-set-member(1)) | As configurações são compatíveis apenas com os suplementos de conteúdo |
| [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) |*Nenhum.*|

Em geral, você usa a API Comum para fazer algo que não é compatível com a API específica do aplicativo. Para obter mais informações sobre como usar a API comum, confira [Modelo do objeto do JavaScript API para Office](../develop/office-javascript-api-object-model.md).

<a name="om-diagram"></a>

## <a name="onenote-object-model-diagram"></a>Diagrama do modelo de objeto do OneNote

O diagrama a seguir representa o que está disponível atualmente na API JavaScript do OneNote.

  ![Diagrama do modelo de objeto do OneNote.](../images/onenote-om.png)

## <a name="see-also"></a>Confira também

- [Desenvolvimento de Suplementos do Office ](../develop/develop-overview.md)
- [Saiba mais sobre Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Criar seu primeiro suplemento do OneNote](../quickstarts/onenote-quickstart.md)
- [Referência da API JavaScript do OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
