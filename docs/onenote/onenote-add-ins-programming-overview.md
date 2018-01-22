# <a name="onenote-javascript-api-programming-overview"></a>Visão geral da programação da API JavaScript do OneNote

O OneNote introduz uma API JavaScript para os suplementos do OneNote Online. Você pode criar suplementos de painel de tarefas e de conteúdo e comandos de suplemento que interagem com objetos do OneNote e conectam-se a serviços Web ou a outros recursos baseados na Web.

>
  **Observação:** Caso pretenda [publicar](../publish/publish.md) o suplemento na Office Store depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação da Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) e a [Página de hospedagem e disponibilidade do suplemento do Office](https://dev.office.com/add-in-availability)).

## <a name="components-of-an-office-add-in"></a>Componentes de um suplemento do Office

Os suplementos consistem de dois componentes básicos:

- Um **aplicativo Web** consiste em uma página da Web e em JavaScript, CSS ou outros arquivos necessários. Estes arquivos podem ser hospedados em qualquer servidor Web ou serviço de hospedagem na Web, como o Microsoft Azure. No OneNote Online, o aplicativo Web exibe um controle de navegação ou iframe.
    
- Um **manifesto XML** que especifica a URL da página da Web do suplemento e os requisitos de acesso, as configurações e os recursos para o suplemento. Este arquivo é armazenado no cliente. Os suplementos do OneNote usam o mesmo formato de [manifesto](https://dev.office.com/docs/add-ins/overview/add-in-manifests) como outros suplementos do Office.

**Suplemento do Office = manifesto + página da Web**

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

    > **Observação**: Chamadas de método para a API (como `context.application.getActiveSection().pages;`) também são adicionadas a fila.

4. Chame **context.sync** para executar todos os comandos na fila na ordem em que eles estão. Isso sincroniza o estado entre o momento em que os scripts e os objetos reais estão sendo executados, além de recuperar as propriedades dos objetos do OneNote carregados para uso no script. Você pode usar o objeto de promessa retornado para o encadeamento de ações adicionais.

Por exemplo: 

```
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

Você pode encontrar objetos do OneNote e operações compatíveis na [Referência API](http://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).

### <a name="accessing-the-common-api-through-the-document-object"></a>Acessar a API comum por meio do objeto *Document*

Use o objeto **Document** para acessar a API comum, como os métodos [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) e [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync). 

Por exemplo:  

```
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
| [Office.context.document.getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) | Apenas **Office.CoercionType.Text** e **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142145.aspx) | Apenas **Office.CoercionType.Text**, **Office.CoercionType.Image** e **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(nome);](https://msdn.microsoft.com/en-us/library/office/fp142180.aspx) | As configurações são compatíveis apenas com os suplementos de conteúdo | 
| [Office.context.document.settings.set(nome, valor);](https://msdn.microsoft.com/en-us/library/office/fp161063.aspx) | As configurações são compatíveis apenas com os suplementos de conteúdo | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

Em geral, você só pode usar a API comum para fazer algo que não seja compatível com a API avançada. Para saber mais sobre como usar a API comum, confira a [documentação](https://dev.office.com/docs/add-ins/overview/office-add-ins) e a [referência](https://dev.office.com/reference/add-ins/javascript-api-for-office) dos suplementos do Office.


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>Diagrama do modelo de objeto do OneNote 
O diagrama a seguir representa o que está disponível atualmente na API JavaScript do OneNote.

  ![Diagrama do modelo de objeto do OneNote](../images/onenote-om.png)


## <a name="additional-resources"></a>Recursos adicionais

- [Crie seu primeiro suplemento do OneNote](onenote-add-ins-getting-started.md)
- [Referência da API JavaScript do OneNote](http://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma de Suplementos do Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
