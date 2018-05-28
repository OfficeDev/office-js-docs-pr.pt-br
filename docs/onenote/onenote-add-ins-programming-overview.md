---
title: Vis?o geral da programa??o da API JavaScript do OneNote
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: aded1210abc11a80c6200a207d3896df8ef4218b
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="onenote-javascript-api-programming-overview"></a>Vis?o geral da programa??o da API JavaScript do OneNote

O OneNote introduz uma API JavaScript para os suplementos do OneNote Online. Voc? pode criar suplementos de painel de tarefas e de conte?do e comandos de suplemento que interagem com objetos do OneNote e conectam-se a servi?os Web ou a outros recursos baseados na Web.

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experi?ncia do Office depois de cri?-lo, verifique se voc? est? em conformidade com as [Pol?ticas de valida??o do AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Por exemplo, para passar na valida??o, seu suplemento deve funcionar em todas as plataformas com suporte aos m?todos que voc? definir (para mais informa??es, confira a [se??o 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [P?gina de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).

## <a name="components-of-an-office-add-in"></a>Componentes de um suplemento do Office

Os suplementos consistem de dois componentes b?sicos:

- Um **aplicativo Web** consiste em uma p?gina da Web e em JavaScript, CSS ou outros arquivos necess?rios. Estes arquivos podem ser hospedados em qualquer servidor Web ou servi?o de hospedagem na Web, como o Microsoft Azure. No OneNote Online, o aplicativo Web exibe um controle de navega??o ou iframe.
    
- Um **manifesto XML** que especifica a URL da p?gina da Web do suplemento e os requisitos de acesso, as configura??es e os recursos para o suplemento. Este arquivo ? armazenado no cliente. Os suplementos do OneNote usam o mesmo formato de [manifesto](../develop/add-in-manifests.md) como outros suplementos do Office.

**Suplemento do Office = manifesto + p?gina da Web**

![Um suplemento do Office consiste em um manifesto e uma p?gina da Web](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>Usar a API JavaScript

Os suplementos usam o contexto de tempo de execu??o do aplicativo host para acessar a API JavaScript. A API tem duas camadas: 

- Uma **API avan?ada** para opera??es espec?ficas do OneNote, acessada por meio do objeto **Application**.
- Uma **API comum** compartilhada entre os aplicativos do Office, acessada por meio do objeto **Document**.

### <a name="accessing-the-rich-api-through-the-application-object"></a>Acessar uma API avan?ada por meio do objeto *Application*.

Use o objeto **Application** para acessar os objetos do OneNote, como **Notebook**, **Section** e **Page**. Com as APIs avan?adas, voc? executa opera??es em lotes em objetos proxy. O fluxo b?sico ser? semelhante a: 

1. Obtenha a inst?ncia do aplicativo do contexto.

2. Crie um proxy que representa o objeto do OneNote com o qual voc? deseja trabalhar. Voc? interage com sincronia com os objetos proxy lendo e gravar suas propriedades e chamando seus m?todos. 

3. Chame **load** no proxy para preench?-lo com valores de propriedade especificados no par?metro. Essa chamada ? adicionada ? fila de comandos.

   > [!NOTE]
   > Chamadas de m?todo para a API (como `context.application.getActiveSection().pages;`) tamb?m s?o adicionadas ? fila.

4. Chame **context.sync** para executar todos os comandos na fila na ordem em que eles est?o. Isso sincroniza o estado entre o momento em que os scripts e os objetos reais est?o sendo executados, al?m de recuperar as propriedades dos objetos do OneNote carregados para uso no seu script. Voc? pode usar o objeto promessa retornado para o encadeamento a??es adicionais.

Por exemplo: 

```js
function getPagesInSection() {
    OneNote.run(function (context) {
        
        // Get the pages in the current section.
        var pages = context.application.getActiveSection().pages;
        
        // Queue a command to load the id and title for each page.            
        pages.load('id,title');
        
        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function () {
                
                // Read the id and title of each page. 
                $.each(pages.items, function(index, page) {
                    var pageId = page.id;
                    var pageTitle = page.title;
                    console.log(pageTitle + ': ' + pageId); 
                });
            })
            .catch(function (error) {
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    });
}
```

Voc? pode encontrar objetos do OneNote e opera??es compat?veis na [Refer?ncia API](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).

### <a name="accessing-the-common-api-through-the-document-object"></a>Acessar a API comum por meio do objeto *Document*

Use o objeto **Document** para acessar a API comum, como os m?todos [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) e [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync). 

Por exemplo:  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```
Os suplementos do OneNote s?o compat?veis apenas com as seguintes APIs comuns:

| API | Observa??es |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) | Apenas **Office.CoercionType.Text** e **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) | Apenas **Office.CoercionType.Text**, **Office.CoercionType.Image** e **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(nome);](https://dev.office.com/reference/add-ins/shared/settings.get) | As configura??es s?o compat?veis apenas com os suplementos de conte?do | 
| [Office.context.document.settings.set(nome, valor);](https://dev.office.com/reference/add-ins/shared/settings.set) | As configura??es s?o compat?veis apenas com os suplementos de conte?do | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

Em geral, voc? s? pode usar a API comum para fazer algo que n?o seja compat?vel com a API avan?ada. Para saber mais sobre como usar a API comum, confira os suplementos do Office [documenta??o](../overview/office-add-ins.md) e [refer?ncia](https://dev.office.com/reference/add-ins/javascript-api-for-office).


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>Diagrama do modelo de objeto do OneNote 
O diagrama a seguir representa o que est? dispon?vel atualmente na API JavaScript do OneNote.

  ![Diagrama do modelo de objeto do OneNote](../images/onenote-om.png)


## <a name="see-also"></a>Veja tamb?m

- [Criar seu primeiro suplemento do OneNote](onenote-add-ins-getting-started.md)
- [Refer?ncia da API JavaScript do OneNote](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vis?o geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
