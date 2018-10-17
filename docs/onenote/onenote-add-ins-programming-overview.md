---
title: Visão geral da programação da API JavaScript do OneNote
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 557fd1807d860960e7d34587d8ad685c15a883fb
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506270"
---
# <a name="onenote-javascript-api-programming-overview"></a>Visão geral da programação da API JavaScript do OneNote

O OneNote introduz uma API JavaScript para os suplementos do OneNote Online. Você pode criar suplementos de painel de tarefas, de conteúdo e de comandos de que interagem com objetos do OneNote e conectam-se a serviços web ou a outros recursos baseados na web.

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) seu suplemento no AppSource e disponibilizá-lo na experiência do Office, verifique se está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deverá funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).

## <a name="components-of-an-office-add-in"></a>Componentes de um suplemento do Office

Os suplementos consistem de dois componentes básicos:

- Um **aplicativo web** consiste em uma página da web e em JavaScript, CSS ou outros arquivos necessários. Estes arquivos podem ser hospedados em qualquer servidor web ou serviço de hospedagem Web, como o Microsoft Azure. No OneNote Online, o aplicativo da Web exibe um controle de navegação ou iframe.
    
- Um **manifesto XML** que especifica a URL da página da web do suplemento e os requisitos de acesso, as configurações e os recursos para o suplemento. Este arquivo é armazenado no cliente. Os suplementos do OneNote usam o mesmo formato de [manifesto](../develop/add-in-manifests.md) como outros suplementos do Office.

**Suplemento do Office = manifesto + página da web**

![Um suplemento do Office consiste em um manifesto e uma página da Web](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>Usar a API JavaScript

Os suplementos usam o contexto de tempo de execução do aplicativo host para acessar a API JavaScript. A API tem duas camadas: 

- Uma **API avançada** para operações específicas do OneNote, acessada por meio do objeto **Application**.
- Uma **API comum** compartilhada entre os aplicativos do Office, acessada por meio do objeto **Document**.

### <a name="accessing-the-rich-api-through-the-application-object"></a>Acessar uma API avançada por meio do objeto *Application*.

Use o objeto **Application** para acessar os objetos do OneNote, como **Notebook**, **Section** e **Page**. Com as APIs avançadas, você executa operações em lotes em objetos proxy. O fluxo básico será semelhante a: 

1. Obtenha a instância do aplicativo do contexto.

2. Crie um proxy que representa o objeto do OneNote com o qual você deseja trabalhar. Você interage com sincronia com os objetos proxy lendo e gravar suas propriedades e chamando seus métodos. 

3. Chame **load** no proxy para preenchê-lo com valores de propriedade especificados no parâmetro. Essa chamada é adicionada à fila de comandos.

   > [!NOTE]
   > Chamadas de método para a API (como `context.application.getActiveSection().pages;`) também são adicionadas à fila.

4. Chame **context.sync** para executar todos os comandos na fila na ordem em que eles estão. Isso sincroniza o estado entre o momento em que os scripts e os objetos reais estão sendo executados, além de recuperar as propriedades dos objetos do OneNote carregados para uso no seu script. Você pode usar o objeto promessa retornado para o encadeamento ações adicionais.

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

Você pode encontrar objetos do OneNote e operações compatíveis na [Referência API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js).

### <a name="accessing-the-common-api-through-the-document-object"></a>Acessar a API comum por meio do objeto *Document*

Use o objeto **Document** para acessar a API comum, como os métodos [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) e [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-). 


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
Os suplementos do OneNote são compatíveis apenas com as seguintes APIs comuns:

| API | Observações |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) | Apenas **Office.CoercionType.Text** e **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) | Apenas **Office.CoercionType.Text**, **Office.CoercionType.Image** e **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(name);](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-) | As configurações são compatíveis apenas com os suplementos de conteúdo | 
| [Office.context.document.settings.set(name, value);](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) | As configurações são compatíveis apenas com os suplementos de conteúdo | 
| [Office.EventType.DocumentSelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) ||

Em geral, você só pode usar a API comum para fazer algo que não seja compatível com a API avançada. Para saber mais sobre como usar a API comum, confira a [documentação](../overview/office-add-ins.md) e a [referência](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) dos suplementos do Office.


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>Diagrama do modelo de objeto do OneNote 
O diagrama a seguir representa o que está disponível atualmente na API JavaScript do OneNote.

  ![Diagrama do modelo de objeto do OneNote](../images/onenote-om.png)


## <a name="see-also"></a>Confira também

- [Criar seu primeiro suplemento do OneNote](onenote-add-ins-getting-started.md)
- [Referência da API JavaScript do OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma de suplementos do Office](../overview/office-add-ins.md)
