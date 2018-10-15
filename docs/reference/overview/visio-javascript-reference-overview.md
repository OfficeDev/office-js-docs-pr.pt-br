# <a name="visio-javascript-api-overview"></a>Visão geral da API JavaScript do Visio

Você pode usar as APIs JavaScript do Visio para incorporar os diagramas do Visio no SharePoint Online. Um diagrama do Visio incorporado é um diagrama que é armazenado em uma biblioteca de documentos do SharePoint e exibido em uma página do SharePoint. Para incorporar um diagrama do Visio, exibia-o em um HTML  elemento  `<iframe>`. Em seguida, você pode usar as APIs JavaScript do Visio para trabalhar programaticamente com o diagrama incorporado.

![Diagrama do Visio em um iframe na página do SharePoint junto com a web part do editor de script](../images/visio-api-block-diagram.png)


Você pode usar as APIs JavaScript do Visio para:

* Interagir com os elementos do diagrama do Visio, como páginas e formas
* Criar uma marcação visual na tela do diagrama do Visio
* Gravar manipuladores personalizados para eventos com o mouse no desenho
* Expor dados do diagrama, como texto da forma, dados da forma e hiperlinks, em sua solução.

Este artigo descreve como usar as APIs JavaScript do Visio com o Visio Online para desenvolver suas soluções para o SharePoint Online. Ele apresenta os principais conceitos que são fundamentais para o uso das APIs, como **EmbeddedSession**, **RequestContext**, e dos objetos proxy do JavaScript, além dos métodos **sync()**, **Visio.run()** e **load()**. Os exemplos de código mostram como aplicar esses conceitos.

## <a name="embeddedsession"></a>EmbeddedSession

O objeto EmbeddedSession inicializa a comunicação entre o quadro do desenvolvedor e o quadro do Visio Online.

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a>Visio.run(session, function(context) { batch })

**Visio.Run()** executa um script em lotes que realiza ações no modelo de objeto do Visio. Os comandos em lotes incluem definições de objetos proxy JavaScript locais e métodos **sync()** que sincronizam o estado entre objetos locais e do Visio e prometem resolução. A vantagem das solicitações em lote no **Visio.run()** é que, quando a promessa é resolvida, todos os objetos de página rastreados que foram alocados durante a execução serão liberados automaticamente.

O método run recebe a sessão e o objeto RequestContext e retorna uma promessa (normalmente, apenas o resultado de **context.sync()**). É possível executar a operação em lotes fora do **Visio.run()**. No entanto, em tal cenário, qualquer referência de objeto de página precisa ser controlada e gerenciada manualmente.

## <a name="requestcontext"></a>RequestContext

O objeto RequestContext facilita solicitações para o aplicativo Visio. Como o quadro do desenvolvedor e o aplicativo Visio Online são executados em dois iframes diferentes, o objeto RequestContext (contexto no próximo exemplo) é necessário para obter acesso ao Visio e aos objetos relacionados, como páginas e formas, do quadro do desenvolvedor.

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a>Objetos proxy

Os objetos JavaScript do Visio declarados e usados em um suplemento são objetos de proxy para os objetos reais em um documento do Visio. Todas as ações executadas em objetos de proxy não são realizadas no Visio, e o estado do documento do Visio não é realizado nos objetos de proxy até que o estado do documento seja sincronizado. O estado do documento é sincronizado quando `context.sync()` é executado.

Por exemplo, o objeto JavaScript local getActivePage é declarado para referenciar a página selecionada. Isso pode ser usado para enfileirar a configuração de suas propriedades e invocar métodos. As ações em tais objetos não são realizadas até que o método **sync()** seja executado.

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a>sync()

O método **sync()** sincroniza o estado entre objetos proxy JavaScript e objetos reais no Visio executando instruções enfileiradas no contexto e recuperando propriedades de objetos do Office carregados para uso em seu código. Esse método retorna uma promessa, que é resolvida quando a sincronização é concluída. 

## <a name="load"></a>load()

O método **load()** é usado para preencher os objetos proxy criados na camada JavaScript do suplemento. Ao tentar recuperar um objeto, como um documento, um objeto proxy local é criado inicialmente na camada JavaScript. Você pode usar esse objeto para colocar a configuração das respectivas propriedades em fila e para invocar métodos. No entanto, você deve invocar inicialmente os métodos **load()** e **sync()** para as propriedades do objeto de leitura ou relações. O método load() é realizado nas propriedades e relações que devem ser carregadas quando você chama o método **sync()**.

O código a seguir mostra a sintaxe para o método **load()**.

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. **properties** é a lista de nomes de propriedades a serem carregados, especificados como sequências de caracteres delimitadas por vírgula ou uma matriz de nomes. Consulte os métodos de **.load()** em cada objeto para obter detalhes.

2. **loadOption** especifica um objeto que descreve as opções selection, expansion, top e skip. Confira as [opções](/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.

## <a name="example-printing-all-shapes-text-in-active-page"></a>Exemplo: Imprimindo todas as formas de texto na página ativa

O exemplo a seguir mostra como imprimir o valor do texto da forma de um objeto de matriz de formas. O método  **Visio.run()** contém um lote de instruções. Como parte desse lote, é criado um objeto proxy que faz referência a formas no documento ativo.

Todos esses comandos são enfileirados e executados quando **context.sync()** é chamado. O método **sync()** retorna uma promessa que pode ser usada para encadeá-lo com outras operações.

```js
Visio.run(session, function (context) {
    var page = context.document.getActivePage();
    var shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++) {
            var shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a>Mensagens de erro

O sistema retorna erros usando um objeto error composto por um código e uma mensagem. A tabela a seguir fornece uma lista de possíveis condições de erro que podem ocorrer.

| error.code            | error.message |
|-----------------------|----------------------------------------------------------------|
| InvalidArgument       | O argumento é inválido, ausente ou tem um formato incorreto. |
| GeneralException      | Ocorreu um erro interno ao processar a solicitação. |
| NotImplemented        | O recurso solicitado não foi implementado.  |
| UnsupportedOperation  | Não há suporte para a operação. |
| AccessDenied          | Você não pode realizar a operação solicitada. |
| ItemNotFound          | O recurso solicitado não existe. |

## <a name="get-started"></a>Introdução

Você pode usar o exemplo nesta seção para começar. Este exemplo mostra como exibir programaticamente o texto da forma selecionada em um diagrama do Visio. Para começar, crie uma página clássica no SharePoint Online ou edite uma página existente. Adicione uma web part do editor de scripts na página e copie e cole o código a seguir.

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        var page = context.document.getActivePage();
        var shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(var i=0; i<shapes.items.length;i++) {
                var shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

Depois disso, você só precisa da URL de um diagrama do Visio com o qual deseja trabalhar. Basta carregar o diagrama do Visio no SharePoint Online e abri-lo no Visio Online. A partir daí, abra a caixa de diálogo "Incorporar" e use a URL de incorporação no exemplo acima.

![Copie a URL do arquivo do Visio da caixa de diálogo "Incorporar"](../images/Visio-embed-url.png)

Se você estiver usando o Visio Online no modo de edição, abra a caixa de diálogo Incorporar navegando até **Arquivo** > **Compartilhamento** > **Incorporar**. Se você estiver usando o Visio Online no modo de exibição, abra a caixa de diálogo "Incorporar" escolhendo '...' e, em seguida, **Incorporar**.

## <a name="open-api-specifications"></a>Especificações abertas da API

À medida que criamos e desenvolvemos novas APIs, elas ficam disponíveis na nossa página [Especificações abertas da API](../openspec.md) para receber seus comentários. Descubra quais novos recursos estão no pipeline e forneça comentários sobre nossas especificações de design.

## <a name="visio-javascript-api-reference"></a>Referência da API JavaScript do Visio

Para obter informações detalhadas sobre a API JavaScript do Visio , consulte a [documentação de referência da API JavaScript do Visio](/javascript/api/visio).
