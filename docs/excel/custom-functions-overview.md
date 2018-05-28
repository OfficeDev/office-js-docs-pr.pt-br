# <a name="create-custom-functions-in-excel-preview"></a>Criar fun??es personalizadas no Excel (Visualiza??o)

Fun??es personalizadas (semelhantes a fun??es definidas pelo usu?rio ou UDFs) permitem que os desenvolvedores adicionem qualquer fun??o JavaScript no Excel usando um suplemento. Os usu?rios podem acessar fun??es personalizadas como qualquer outra fun??o nativa no Excel (como `=SUM()`). Este artigo explica como criar fun??es personalizadas no Excel.

A ilustra??o a seguir mostra como um usu?rio final pode inserir uma fun??o personalizada em uma c?lula. A fun??o que adiciona 42 a um par de n?meros.

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

Aqui est? o c?digo para a mesma fun??o personalizada.

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

As fun??es personalizadas agora est?o dispon?veis no Developer Preview para Windows, Mac e Excel Online. Siga estas etapas para experiment?-las:

1.  Instale o Office (compila??o 9325 no Windows ou 13.329 no Mac) e participe do programa [Office Insider](https://products.office.com/en-us/office-insider). (Observe que n?o ? suficiente apenas obter a compila??o mais recente; o recurso ser? desabilitado em qualquer compila??o at? voc? ingressar no programa Insider)
2.  Clone o reposit?rio [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) e siga as instru??es no README.md para iniciar o suplemento no Excel, fazer altera??es no c?digo e depurar.
3.  Digite `=CONTOSO.ADD42(1,2)` em qualquer c?lula e pressione **Inserir** para executar a fun??o personalizada.

Confira a se??o **Problemas conhecidos** no final deste artigo, que inclui as limita??es atuais das fun??es personalizadas e que ser? atualizado com o tempo.

## <a name="learn-the-basics"></a>No??es b?sicas

No reposit?rio de exemplo clonado, voc? ver? os seguintes arquivos:

- **customfunctions.js**, que cont?m o c?digo de fun??o personalizado (veja o exemplo de c?digo simples acima para a `ADD42` fun??o).
- **customfunctions.json**, que cont?m o registro JSON que informa ao Excel sobre sua fun??o personalizada. O registro faz com que suas fun??es personalizadas apare?am na lista de fun??es dispon?veis exibidas quando um usu?rio digita em uma c?lula.
- **customfunctions.html**, que fornece um &lt;Script&gt; de refer?ncia para o arquivo JS. Este arquivo n?o ? exibido na interface do usu?rio do Excel.
- **customfunctions.xml**, que informa ao Excel a localiza??o dos arquivos HTML, JavaScript e JSON; e tamb?m especifica um namespace para todas as fun??es personalizadas instaladas com o suplemento.

### <a name="json-file-customfunctionsjson"></a>Arquivo JSON (customfunctions.json)

O c?digo a seguir em customfunctions.json especifica os metadados para a mesma fun??o `ADD42`.

> [!NOTE]
> Informa??es de refer?ncia detalhadas para o arquivo JSON, incluindo op??es n?o usadas neste exemplo, est?o em [Registro de Fun??es Personalizadas JSON](https://dev.office.com/reference/add-ins/custom-functions-json).

Observe que, para este exemplo:

- H? apenas uma fun??o personalizada, portanto, h? apenas um membro da `functions` matriz.
- A propriedade `name` define o nome da fun??o. Como voc? viu no gif animado mostrado anteriormente, um namespace (`CONTOSO`) ? anexado ao nome da fun??o no menu de preenchimento autom?tico do Excel. Esse prefixo ? definido no manifesto do suplemento, descrito abaixo. O prefixo e o nome da fun??o s?o separados por um ponto e, por conven??o, nomes de fun??o e prefixos s?o mai?sculos. Para usar a fun??o personalizada, o usu?rio digita o namespace seguido pelo nome da fun??o (`ADD42`) em uma c?lula, neste caso `=CONTOSO.ADD42`. O prefixo deve ser usado como identificador para a sua empresa ou para o suplemento. 
- O `description` aparece no menu de preenchimento autom?tico do Excel.
- Quando o usu?rio solicitar ajuda para uma fun??o, o Excel abre um painel de tarefas e exibe a p?gina da Web encontrada no URL especificado em `helpUrl` .
- A propriedade `result` especifica o tipo de informa??o retornada pela fun??o para o Excel. A propriedade filho `type` pode `"string"`, `"number"`ou `"boolean"`. A propriedade `dimensionality` pode ser `scalar` ou `matrix` (uma matriz bidimensional de valores do `type` especificado.)
- A matriz `parameters` especifica, *em ordenar*, o tipo de dado em cada par?metro que ? passado para a fun??o. As propriedades filho `name` e `description` s?o usadas no intellisense do Excel. As propriedades filho `type` e `dimensionality` s?o id?nticas ?s propriedades filho da propriedade `result` descrita acima.
- A propriedade `options` permite que voc? personalize alguns aspectos de como e quando o Excel executa a fun??o. H? mais informa??es sobre essas op??es posteriormente neste artigo.

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
> As fun??es personalizadas s?o registradas quando um usu?rio executa o suplemento pela primeira vez. Depois disso, eles estar?o dispon?veis, para o mesmo usu?rio, em todas as pastas de trabalho (n?o apenas naquela em que o suplemento foi executado inicialmente).

As configura??es do servidor para o arquivo JSON devem ter [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) habilitado para que fun??es personalizadas funcionem corretamente no Excel Online.


### <a name="manifest-file-customfunctionsxml"></a>Arquivo de manifesto (customfunctions.xml)


O seguinte ? um exemplo da marca??o `<ExtensionPoint>` e `<Resources>` que voc? inclui no manifesto do suplemento para permitir que o Excel execute suas fun??es. Observe o seguinte sobre essa marca??o:

- O elemento `<Script>` e a identifica??o do recurso correspondente especificam a localiza??o do arquivo JavaScript com suas fun??es.
- O elemento `<Page>` e a identifica??o do recurso correspondente especificam a localiza??o da p?gina HTML do seu suplemento. A p?gina HTML inclui uma marca `<Script>` que carrega o arquivo JavaScript (customfunctions.js). A p?gina HTML ? uma p?gina oculta e nunca ? exibida na interface de usu?rio.
- O elemento `<Metadata>` e a identifica??o do recurso correspondente especificam a localiza??o do arquivo JSON.
- Um elemento `<Namespace>` e a identifica??o do recurso correspondente especificam o prefixo para todas as fun??es personalizadas no suplemento.


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

## <a name="initializing-custom-functions"></a>Inicializa??o de fun??es personalizadas

Seu c?digo deve inicializar o recurso de fun??es personalizadas antes de us?-lo. Voc? pode fazer isso em uma marca de &lt;Script&gt; no arquivo HTML (customfunctions.html) ou na parte superior do arquivo JavaScript (customfunctions.js). Na visualiza??o de fun??es personalizadas, voc? pode escolher entre duas sintaxes para a inicializa??o. O arquivo HTML no reposit?rio usa a seguinte sintaxe:

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

Voc? tamb?m pode usar a seguinte sintaxe:

```js
Office.Preview.StartCustomFunctions();
```

## <a name="synchronous-and-asynchronous-functions"></a>Fun??es s?ncronas e ass?ncronas

A fun??o `ADD42` acima ? s?ncrona em rela??o ao Excel (designada pela configura??o da op??o `"sync": true` no arquivo JSON). As fun??es s?ncronas oferecem desempenho r?pido porque s?o executadas no mesmo processo que o Excel e em paralelo durante o c?lculo multithreaded.   

Por outro lado, se sua fun??o personalizada recupera dados da Web, ela dever? ser ass?ncrona em rela??o ao Excel. Fun??es ass?ncronas devem:

1. Retornar uma Promise do JavaScript para o Excel.
3. Resolver a Promise com o valor final usando a fun??o de retorno de chamada.

O c?digo a seguir mostra um exemplo de uma fun??o ass?ncrona que recupera a temperatura de um term?metro. Observe que `sendWebRequest` ? uma fun??o hipot?tica, n?o especificada aqui, que usa o XHR para chamar um servi?o da Web de temperatura.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

Fun??es ass?ncronas exibem um erro tempor?rio `GETTING_DATA` na c?lula enquanto o Excel aguarda o resultado final. Os usu?rios podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.

> [!NOTE]
> Fun??es personalizadas s?o ass?ncronas por padr?o. Para designar fun??es como s?ncronas, defina a op??o `"sync": true` na propriedade `options` para a fun??o personalizada no arquivo JSON de registro.

## <a name="streamed-functions"></a>Fun??es de fluxo cont?nuo

Uma fun??o ass?ncrona pode ser de fluxo cont?nuo. Fun??es personalizadas de fluxo cont?nuo permitem que voc? insira dados em c?lulas repetidamente ao longo do tempo, sem precisar esperar que o Excel ou os usu?rios solicitem rec?lculos. O exemplo a seguir ? uma fun??o personalizada que adiciona um n?mero ao resultado a cada segundo. Observe o seguinte sobre este c?digo:

- O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.
- O par?metro final, `caller`, nunca ? especificado no c?digo de registro e nunca ? exibido no menu de preenchimento autom?tico para usu?rios do Excel ao inserir a fun??o. ? a fun??o de retorno de chamada `setResult` usada para passar dados da fun??o para o Excel para atualizar o valor de uma c?lula.
- Para que o Excel passe a fun??o `setResult` no objeto `caller`, voc? deve declarar suporte para fluxo cont?nuo durante o registro da fun??o, definindo a op??o `"stream": true` na propriedade `options` para a fun??o personalizada no arquivo JSON de registro.

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

Voc? pode cancelar fun??es e fun??es ass?ncronas de streaming. ? importante cancelar as chamadas de fun??o para reduzir o consumo de largura de banda, a mem?ria de trabalho e a carga da CPU. O Excel cancela chamadas de fun??es nas seguintes situa??es:

- O usu?rio edita ou exclui uma c?lula que faz refer?ncia ? fun??o.
- ? alterado um dos argumentos (entradas) para a fun??o. Nesse caso, uma nova chamada de fun??o ? disparada, al?m do cancelamento.
- O usu?rio aciona um rec?lculo manualmente. Como no caso acima, uma nova chamada de fun??o ? disparada, al?m do cancelamento.

Voc? *deve* implementar um manipulador de cancelamento para todas as fun??es de fluxo cont?nuo. Fun??es ass?ncronas e que n?o sejam de fluxo cont?nuo podem ou n?o ser cancel?veis; a decis?o ? sua. Fun??es s?ncronas n?o podem ser canceladas.

Para tornar uma fun??o cancel?vel, defina a op??o `"cancelable": true` na propriedade `options` para a fun??o personalizada no arquivo JSON de registro.

O c?digo a seguir mostra o exemplo anterior com o cancelamento implementado. No c?digo, o objeto `caller` cont?m uma fun??o `onCanceled` que deve ser definida para cada fun??o personalizada cancel?vel.

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

Fun??es personalizadas ass?ncronas podem salvar dados em vari?veis JavaScript globais. Em chamadas subsequentes, sua fun??o personalizada pode usar os valores salvos nessas vari?veis. O estado salvo ? ?til quando os usu?rios adicionam a mesma fun??o personalizada a mais de uma c?lula, porque todas as inst?ncias da fun??o podem compartilhar o estado. Por exemplo, voc? pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.

O c?digo a seguir mostra uma implementa??o da fun??o de fluxo cont?nuo anterior de temperatura que salva o estado de forma global. Observe o seguinte sobre este c?digo:

- `refreshTemperature` ? uma fun??o de fluxo cont?nuo que l? a temperatura de um determinado term?metro a cada segundo. Novas temperaturas s?o salvas na vari?vel `savedTemperatures`, mas o valor da c?lula n?o ? atualizado diretamente. N?o deve ser chamada diretamente de uma c?lula da planilha, *por isso n?o est? registrada no arquivo JSON*.
- `streamTemperature` atualiza os valores de temperatura exibidos na c?lula a cada segundo e usa vari?vel `savedTemperatures` como fonte de dados. Deve ser registrada no arquivo JSON e nomeada com todas as letras mai?sculas, `STREAMTEMPERATURE`.
- Os usu?rios podem chamar `streamTemperature` de v?rias c?lulas na interface de usu?rio do Excel. Cada chamada l? dados da mesma vari?vel `savedTemperatures`.

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
> Fun??es s?ncronas (designadas pela configura??o da op??o `"sync": true` no arquivo JSON) n?o podem compartilhar estado porque o Excel faz o paralelismo delas durante o c?lculo multithreaded. Somente fun??es ass?ncronas podem compartilhar estado porque as fun??es s?ncronas de um suplemento compartilham o mesmo contexto JavaScript em cada sess?o.

## <a name="working-with-ranges-of-data"></a>Trabalhar com intervalos de dados

Sua fun??o personalizada pode levar a um intervalo de dados como um par?metro ou voc? pode retornar um intervalo de dados de uma fun??o personalizada.

Por exemplo, suponha que sua fun??o retorne o segundo maior valor de um intervalo de n?meros armazenados no Excel. A fun??o a seguir usa o par?metro `values`, que ? um tipo de par?metro `Excel.CustomFunctionDimensionality.matrix`. Note que no registro JSON para esta fun??o, voc? definiria a propriedade `type` do par?metro para `matrix`.

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

Como pode ver, os intervalos s?o tratados em JavaScript como matrizes de matrizes de linhas (como uma matriz bidimensional).

## <a name="known-issues"></a>Problemas conhecidos

- Descri??es de par?metro e URLs de Ajuda ainda n?o s?o usados pelo Excel.
- Fun??es personalizadas n?o est?o atualmente dispon?veis no Excel para clientes m?veis.
- Atualmente, os suplementos dependem de um processo de navegador oculto para executar fun??es personalizadas ass?ncronas. No futuro, o JavaScript ser? executado diretamente em algumas plataformas para garantir que as fun??es personalizadas sejam mais r?pidas e usem menos mem?ria. Al?m disso, a p?gina HTML referenciada pelo elemento `<Page>` no manifesto n?o ser? necess?ria para a maioria das plataformas, j? que o Excel executa o JavaScript diretamente. Para se preparar para essa altera??o, certifique-se de que suas fun??es personalizadas n?o usem o DOM da p?gina da Web. As APIs de hospedagem suportadas para acessar a Web ser?o [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) e [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) usando GET ou POST.
- Fun??es vol?teis (aquelas que recalculam automaticamente sempre que dados n?o relacionados s?o alterados na planilha) ainda n?o s?o suportadas.
- A depura??o s? est? habilitada para fun??es ass?ncronas no Excel para Windows.
- A implanta??o por meio do Portal de Administra??o do Office 365 e do AppSource ainda n?o est? habilitada.
- Fun??es personalizadas no Excel Online podem parar de funcionar durante uma sess?o ap?s um per?odo de inatividade. Atualize a p?gina do navegador (F5) e insira novamente uma fun??o personalizada para restaurar o recurso.

## <a name="changelog"></a>Log de mudan?as

- **7 de novembro de 2017**: enviados exemplos e visualiza??es de fun??es personalizadas
- **20 de Nov de 2017**: corre??o de bug de compatibilidade para quem usa as vers?es 8801 e posteriores
- **28 de novembro de 2017**: Enviado o suporte para cancelamento em fun??es ass?ncronas (requer altera??o para fun??es de fluxo cont?nuo)
- **7 de maio de 2018**: enviado o suporte para Mac, Excel Online e fun??es s?ncronas em execu??o no processo
