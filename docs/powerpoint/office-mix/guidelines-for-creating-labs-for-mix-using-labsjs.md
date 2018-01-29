
# <a name="guidelines-for-creating-labs-for-mix-using-labsjs"></a>Diretrizes para criar laboratórios para o Mix usando LabsJS



A biblioteca LabsJS (labs.js) oferece suporte à produção de Suplementos do Office especializados (chamados de laboratórios) que se integram ao Office Mix. Em seguida, o Office Mix processa os laboratórios usando o Microsoft PowerPoint. Embora esses componentes sejam chamados de "laboratórios", vamos esclarecer que o que estamos criando são Suplementos do Office especiais que são Suplementos do Office Mix.

O conteúdo do LabsJS ajuda a implementar a API JavaScript labs.js, fornecendo orientação e exemplos. Esta biblioteca é criada com base na [API JavaScript para Office](http://dev.office.com/reference/add-ins/javascript-api-for-office) (Office.js) e fornece uma camada de abstração que é otimizada para suplementos incorporados no Office Mix.


## <a name="general-guidelines"></a>Diretrizes gerais


Veja a seguir algumas diretrizes gerais para ajudar na produção de suplementos usando a API LabJS.


### <a name="scripts"></a>Scripts

Como a biblioteca labs.js é uma camada de abstração no office.js e, portanto, tem uma dependência do office.js, os arquivos de biblioteca office.js e labs.js devem ser incluídos em seus projetos de desenvolvimento. 

Você pode fazer referência à biblioteca office.js aqui:  `<script src="https://sforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>`.

A biblioteca labs.js está incluída com o SDK da LabsJS. Como alternativa, você pode fazer referência à biblioteca labs.js em uma CDN em <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. Observe que a versão de produção de seu laboratório deve fazer referência à versão armazenada na CDN.


 >**Observação**:  Além do arquivo JavaScript (labs-1.0.4.js), fornecemos um arquivo de definição TypeScript da API de laboratórios (labs-1.0.4.d.ts). O arquivo de definição foi compilado com base no TypeScript versão 0.9.1.1.


### <a name="callbacks-and-error-handling"></a>Retornos de chamada e manipulação de erros

Vários métodos na API labs.js operam de forma assíncrona. Para essas operações, a API adota uma interface padrão de retorno de chamada, **ILabCallback**. 


```js
function(err, result) {
}
```

O método de retorno de chamada usa dois parâmetros, _err_ e _result_. O campo _err_ permanece **null**, a menos que ocorra um erro. O campo _result_ retorna o resultado da operação.

A operação de retorno de chamada nunca é acionada imediatamente, mesmo se o resultado estiver disponível imediatamente. Em vez disso, é acionada em uma execução separada do loop de eventos do JavaScript (por meio da chamada **setTimeout**). Adotando essa definição de retorno de chamada, você pode integrar facilmente a labs.js com sua promessa de API preferida. Por exemplo, você pode substituir as promessas de jQuery para esses retornos de chamada por um método de tradução simples, como mostra o exemplo a seguir.




```js
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### <a name="lab-host-and-defaultlabhost"></a>Host de laboratório e DefaultLabHost

O host de laboratório (**ILabHost**) é o driver subjacente que oferece suporte ao desenvolvimento de laboratórios. Por padrão, ele está definido como um host que é integrado à office.js.

Para fins de teste, e para executar seu laboratório dentro de labhost.html, você precisará trocar para um host que funciona no ambiente de simulação. O exemplo de código a seguir mostra como fazer isso usando um parâmetro de consulta. Como alternativa, você pode alterar **DefaultHostBuilder** para integrar seu suplemento de laboratório com uma plataforma diferente.




```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### <a name="initialization"></a>Inicialização

A inicialização estabelece o caminho de comunicação entre o laboratório e seu host. Inicialize seu laboratório chamando o seguinte.


```js
Labs.connect((err, connectionResponse) => {});
```

Depois de inicializar, você pode chamar outros métodos para a API labs.js. O parâmetro _connectionResponse_ contém informações sobre o host, o usuário e outras informações relacionadas à conexão. Para saber mais sobre os valores retornados, confira [Labs.Core.IConnectionResponse](http://dev.office.com/reference/add-ins/office-mix/labs.core.iconnectionresponse).


### <a name="time-format"></a>Formato de hora

A Labs.js armazena números como milissegundos decorridos desde 1º de janeiro de 1970 UTC. Isso coincide com o formato de data no [objeto Date](http://msdn.microsoft.com/pt-br/library/ie/cd9w2te4%28v=vs.94%29.aspx) JavaScript.


### <a name="timeline"></a>Linha do tempo

O laboratório também pode interagir com o cronograma do reprodutor de lição. A linha do tempo permite que o laboratório peça ao reprodutor de lição que avance para o próximo slide. O objeto de linha do tempo é recuperado chamando o método **Labs.getTimeline**.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="handling-events"></a>Manipulando eventos


A API de eventos LabsJS rastreia eventos específicos ao laboratório e permite que você adicione manipuladores de eventos para que possa responder ou agir sobre os eventos. Os métodos de evento, que são três, estão no objeto **EventTypes**: **ModeChanged**,  **Activate** e **Deactivate**. 


### <a name="mode-change"></a>Alteração de modo

O evento **ModeChanged** é disparado quando laboratório especificado muda do modo de edição para o modo de exibição. O modo de edição fica visível quando o laboratório é exibido no modo de edição do PowerPoint. O modo de exibição fica visível quando o PowerPoint está processando a apresentação de slides, ou quando o laboratório está sendo exibido no reprodutor de lição do Office Mix. O modo de exibição deve sempre exibir o que o usuário vê ao realizar o laboratório. O modo de edição permite que o usuário configure o laboratório.

Os dados no objeto **ModeChangedEventData** que é passado para o retorno de chamada contêm informações sobre o modo atual. O código a seguir mostra como usar o evento **ModeChanged**.




```js
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### <a name="activate"></a>Ativar

O evento **activate** é disparado quando o slide do PowerPoint no qual o laboratório está no momento fica ativo no reprodutor de lição.


```js
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### <a name="deactivate"></a>Desativar

O evento **deactivate** é disparado quando o slide do PowerPoint no qual o laboratório está no momento não é mais o slide ativo.


```js
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### <a name="timeline"></a>Linha do tempo

O laboratório também pode interagir com o cronograma do reprodutor de lição. A linha do tempo permite que o laboratório peça ao reprodutor de lição que avance para o próximo slide. O objeto de linha do tempo é recuperado chamando o método **Labs.getTimeline**.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="additional-resources"></a>Recursos adicionais



- [Suplementos do Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
