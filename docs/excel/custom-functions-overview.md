# <a name="create-custom-functions-in-excel-preview"></a>Criar funções personalizadas no Excel (Visualização)

Funções personalizadas (semelhantes a funções definidas pelo usuário ou UDFs) permitem que os desenvolvedores adicionem qualquer função JavaScript no Excel usando um suplemento. Os usuários então podem acessar funções personalizadas como qualquer outra função nativa do Excel (como `=SUM()`). Este artigo explica como criar funções personalizadas no Excel.

A ilustração a seguir mostra como um usuário final pode inserir uma função personalizada em uma célula. A função que adiciona 42 a um par de números.

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

Aqui está o código para a mesma função personalizada.

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

As funções personalizadas agora estão disponíveis no Developer Preview para Windows, Mac e Excel Online. Siga estas etapas para experimentá-las:

1.  Instale o Office (compilação 9325 no Windows ou 13.329 no Mac) e participe do programa [Office Insider](https://products.office.com/en-us/office-insider). (Observe que não é suficiente apenas obter a compilação mais recente; o recurso será desabilitado em qualquer compilação até você ingressar no programa Insider)
2.  Clone o repositório [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) e siga as instruções no README.md para iniciar o suplemento no Excel, fazer alterações no código e depurar.
3.  Digite `=CONTOSO.ADD42(1,2)` em qualquer célula e pressione **Enter** para executar a função personalizada.

Confira a seção **Problemas conhecidos** no final deste artigo, que inclui as limitações atuais das funções personalizadas e que será atualizado com o tempo.

## <a name="learn-the-basics"></a>Noções básicas

No repositório de exemplo clonado, você verá os seguintes arquivos:

- **customfunctions.js**, que contém o código de função personalizado (veja o exemplo de código simples acima para a `ADD42` função).
- **customfunctions.json**, que contém o registro JSON que informa ao Excel sobre sua função personalizada. O registro faz com que suas funções personalizadas apareçam na lista de funções disponíveis exibidas quando um usuário digita em uma célula.
- **customfunctions.html**, que fornece um &lt;Script&gt; de referência para o arquivo JS. Este arquivo não é exibido na interface do usuário no Excel.
- **customfunctions.xml**, que informa ao Excel a localização dos arquivos HTML, JavaScript e JSON; e também especifica um namespace para todas as funções personalizadas instaladas com o suplemento.

### <a name="json-file-customfunctionsjson"></a>Arquivo JSON (customfunctions.json)

O código a seguir em customfunctions.json especifica os metadados para a mesma função `ADD42`.

> [!NOTE]
> Informações de referência detalhadas para o arquivo JSON, incluindo opções não usadas neste exemplo, estão em [Registro de Funções Personalizadas JSON](custom-functions-json.md).

Observe que, para este exemplo:

- Há apenas uma função personalizada, portanto, há apenas um membro da `functions` matriz.
- A propriedade `name` define o nome da função. Como você viu no gif animado mostrado anteriormente, um namespace (`CONTOSO`) é anexado ao nome da função no menu de preenchimento automático do Excel. Esse prefixo é definido no manifesto do suplemento, descrito abaixo. O prefixo e o nome da função são separados por um ponto e, por convenção, nomes de função e prefixos são maiúsculos. Para usar a função personalizada, o usuário digita o namespace seguido pelo nome da função (`ADD42`) em uma célula, neste caso `=CONTOSO.ADD42`. O prefixo deve ser usado como identificador para a sua empresa ou para o suplemento. 
- O `description`  aparece no menu de preenchimento automático do Excel.
- Quando o usuário solicitar ajuda para uma função, o Excel abre um painel de tarefas e exibe a página da Web encontrada no URL especificado em `helpUrl` .
- A propriedade `result` especifica o tipo de informação retornada pela função para o Excel. A propriedade filho `type` pode `"string"`, `"number"`ou `"boolean"`. A propriedade `dimensionality` pode ser `scalar` ou `matrix` (uma matriz bidimensional de valores do `type` especificado.)
- A matriz `parameters` especifica, *em ordenar*, o tipo de dado em cada parâmetro que é passado para a função. As propriedades filho `name` e `description` são usadas no intellisense do Excel. As propriedades filho `type` e `dimensionality` são idênticas às propriedades filho da propriedade `result` descrita acima.
- A propriedade `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função. Há mais informações sobre essas opções posteriormente neste artigo.

 ```js
{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "name": "ADD42", 
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "number 1",
                    "description": "the first number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                },
                {
                    "name": "number 2",
                    "description": "the second number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": true
            }
        }
    ]
}
```

> [!NOTE]
> As funções personalizadas são registradas quando um usuário executa o suplemento pela primeira vez. Depois disso, eles estarão disponíveis, para o mesmo usuário, em todas as pastas de trabalho (não apenas naquela em que o suplemento foi executado inicialmente).

As configurações do servidor para o arquivo JSON devem ter [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) habilitado para que funções personalizadas funcionem corretamente no Excel Online.


### <a name="manifest-file-customfunctionsxml"></a>Arquivo de manifesto (customfunctions.xml)


O seguinte é um exemplo da marcação `<ExtensionPoint>` e `<Resources>` que você inclui no manifesto do suplemento para permitir que o Excel execute suas funções. Observe o seguinte sobre essa marcação:

- O elemento `<Script>` e a identificação do recurso correspondente especificam a localização do arquivo JavaScript com suas funções.
- O elemento `<Page>` e a identificação do recurso correspondente especificam a localização da página HTML do seu suplemento. A página HTML inclui uma marca `<Script>` que carrega o arquivo JavaScript (customfunctions.js). A página HTML é uma página oculta e nunca é exibida na interface de usuário.
- O elemento `<Metadata>` e a identificação do recurso correspondente especificam a localização do arquivo JSON.
- Um elemento `<Namespace>` e a identificação do recurso correspondente especificam o prefixo para todas as funções personalizadas no suplemento.


```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="residjs" />
                    </Script>
                    <Page>
                        <SourceLocation resid="residhtml"/>
                    </Page>
                    <Metadata>
                        <SourceLocation resid="residjson" />
                    </Metadata>
                    <Namespace resid="residNS" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="residjson" DefaultValue="http://127.0.0.1:8080/customfunctions.json" />
            <bt:Url id="residjs" DefaultValue="http://127.0.0.1:8080/customfunctions.js" />
            <bt:Url id="residhtml" DefaultValue="http://127.0.0.1:8080/customfunctions.html" />
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="residNS" DefaultValue="CONTOSO" />
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>

```

## <a name="initializing-custom-functions"></a>Inicialização de funções personalizadas

Seu código deve inicializar o recurso de funções personalizadas antes de usá-lo. Você pode fazer isso em uma marca de &lt;Script&gt; no arquivo HTML (customfunctions.html) ou na parte superior do arquivo JavaScript (customfunctions.js). Na visualização de funções personalizadas, você pode escolher entre duas sintaxes para a inicialização. O arquivo HTML no repositório usa a seguinte sintaxe:

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

Você também pode usar a seguinte sintaxe:

```js
Office.Preview.StartCustomFunctions();
```

## <a name="synchronous-and-asynchronous-functions"></a>Funções síncronas e assíncronas

A função `ADD42` acima é síncrona em relação ao Excel (designada pela configuração da opção `"sync": true` no arquivo JSON). As funções síncronas oferecem desempenho rápido porque são executadas no mesmo processo que o Excel e em paralelo durante o cálculo multithreaded.   

Por outro lado, se sua função personalizada recupera dados da Web, ela deverá ser assíncrona em relação ao Excel. Funções assíncronas devem:

1. Retornar uma Promise do JavaScript para o Excel.
3. Resolver a Promise com o valor final usando a função de retorno de chamada.

O código a seguir mostra um exemplo de função assíncrona personalizada que recupera a temperatura de um termômetro. Observe que `sendWebRequest` é uma função hipotética, não especificada aqui, que usa o XHR para chamar um serviço da Web de temperatura.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

Funções assíncronas exibem um erro temporário `GETTING_DATA` na célula enquanto o Excel aguarda o resultado final. Os usuários podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.

> [!NOTE]
> Funções personalizadas são assíncronas por padrão. Para designar funções como síncronas, defina a opção `"sync": true` na propriedade `options` para a função personalizada no arquivo JSON de registro.

## <a name="streamed-functions"></a>Funções de transmissão

Uma função assíncrona pode ser de fluxo contínuo. Funções personalizadas de transmissão permitem que você insira dados em células repetidamente ao longo do tempo, sem precisar esperar que o Excel ou os usuários solicitem recálculos. O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo. Observe o seguinte sobre este código:

- O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.
- O parâmetro final, `caller`, nunca é especificado no código de registro e nunca é exibido no menu de preenchimento automático para usuários do Excel ao inserir a função. É a função de retorno de chamada `setResult` usada para passar dados da função para o Excel para atualizar o valor de uma célula.
- Para que o Excel passe a função `setResult` no objeto `caller`, você deve declarar suporte para fluxo contínuo durante o registro da função, definindo a opção `"stream": true` na propriedade `options` para a função personalizada no arquivo JSON de registro.

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a>Cancelamento

Você pode cancelar funções e funções assíncronas de streaming. É importante cancelar as chamadas de função para reduzir o consumo de largura de banda, a memória de trabalho e a carga da CPU. O Excel cancela chamadas de funções nas seguintes situações:

- O usuário edita ou exclui uma célula que faz referência à função.
- É alterado um dos argumentos (entradas) para a função. Nesse caso, uma nova chamada de função é disparada, além do cancelamento.
- O usuário aciona um recálculo manualmente. Como no caso acima, uma nova chamada de função é disparada, além do cancelamento.

Você *deve* implementar um manipulador de cancelamento para todas as funções de fluxo contínuo. Funções assíncronas e que não sejam de fluxo contínuo podem ou não ser canceláveis; a decisão é sua. Funções síncronas não podem ser canceladas.

Para tornar uma função cancelável, defina a opção `"cancelable": true` na propriedade `options` para a função personalizada no arquivo JSON de registro.

O código a seguir mostra o exemplo anterior com o cancelamento implementado. No código, o objeto `caller` contém uma função `onCanceled` que deve ser definida para cada função personalizada cancelável.

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a>Compartilhamento e salvamento de estado

Funções personalizadas assíncronas podem salvar dados em variáveis JavaScript globais. Em chamadas subsequentes, sua função personalizada pode usar valores salvos nessas variáveis. O estado salvo é útil quando os usuários adicionam a mesma função personalizada a mais de uma célula, porque todas as instâncias da função podem compartilhar o estado. Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.

O código a seguir mostra uma implementação da função anterior de fluxo contínuo de temperatura que salva o estado de forma global. Observe o seguinte sobre este código:

- `refreshTemperature` é uma função de transmissão que lê a temperatura de um determinado termômetro a cada segundo. Novas temperaturas são salvas na variável `savedTemperatures`, mas o valor da célula não é atualizado diretamente. Não deve ser chamada diretamente de uma célula da planilha, *por isso não está registrada no arquivo JSON*.
- `streamTemperature` atualiza os valores de temperatura exibidos na célula a cada segundo e usa variável `savedTemperatures` como fonte de dados. Deve ser registrada no arquivo JSON e nomeada com todas as letras maiúsculas, `STREAMTEMPERATURE`.
- Os usuários podem chamar `streamTemperature` de várias células na interface de usuário do Excel. Cada chamada lê dados da mesma variável `savedTemperatures`.

```js
var savedTemperatures;

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequest(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

> [!NOTE]
> Funções síncronas (designadas pela configuração da opção `"sync": true` no arquivo JSON) não podem compartilhar estado porque o Excel faz o paralelismo delas durante o cálculo multithreaded. Somente funções assíncronas podem compartilhar estado porque as funções síncronas de um suplemento compartilham o mesmo contexto JavaScript em cada sessão.

## <a name="working-with-ranges-of-data"></a>Trabalhar com intervalos de dados

Sua função personalizada pode levar a um intervalo de dados como um parâmetro ou você pode retornar um intervalo de dados de uma função personalizada.

Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel. A função a seguir usa o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`. Note que no registro JSON para esta função, você definiria a propriedade `type` do parâmetro para `matrix`.

```js
function secondHighest(values){ 
     var highest = values[0][0], secondHighest = values[0][0];
     for(var i = 0; i < values.length; i++){
         for(var j = 1; j < values[i].length; j++){
             if(values[i][j] >= highest){
                 secondHighest = highest;
                 highest = values[i][j];
             }
             else if(values[i][j] >= secondHighest){
                 secondHighest = values[i][j];
             }
         }
     }
     return secondHighest;
 }
```

Como pode ver, os intervalos são tratados em JavaScript como matrizes de matrizes de linhas (como uma matriz bidimensional).

## <a name="known-issues"></a>Problemas conhecidos

- As descrições de URLs e parâmetros de ajuda ainda não são usadas pelo Excel.
- Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.
- Atualmente, os suplementos dependem de um processo de navegador oculto para executar funções personalizadas assíncronas. No futuro, o JavaScript será executado diretamente em algumas plataformas para garantir que as funções personalizadas sejam mais rápidas e usem menos memória. Além disso, a página HTML referenciada pelo elemento `<Page>` no manifesto não será necessária para a maioria das plataformas, já que o Excel executa o JavaScript diretamente. Para se preparar para essa alteração, certifique-se de que suas funções personalizadas não usem o DOM da página da Web. As APIs de hospedagem suportadas para acessar a Web serão [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) e [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) usando GET ou POST.
- Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não são suportadas.
- A depuração só está habilitada para funções assíncronas no Excel para Windows.
- A implantação por meio do Portal de Administração do Office 365 e do AppSource ainda não está habilitada.
- Funções personalizadas no Excel Online podem parar de funcionar durante uma sessão após um período de inatividade. Atualize a página do navegador (F5) e insira novamente uma função personalizada para restaurar o recurso.

## <a name="changelog"></a>Log de mudanças

- **7 de novembro de 2017**: enviados exemplos e visualizações de funções personalizadas
- **20 de Nov de 2017**: correção de bug de compatibilidade para quem usa as versões 8801 e posteriores
- **28 de novembro de 2017**: enviado o suporte para cancelamento em funções assíncronas (requer a alteração de funções de streaming)
- **7 de maio de 2018**: enviado o suporte para Mac, Excel Online e funções síncronas em execução no processo
