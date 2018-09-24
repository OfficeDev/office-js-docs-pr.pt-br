---
ms.date: 09/20/2018
description: Criar uma função personalizada no Excel usando o JavaScript.
title: Criar funções personalizadas no Excel (Visualização)
ms.openlocfilehash: 295152ca14cf56293d51b8b0512b729373841208
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062127"
---
# <a name="create-custom-functions-in-excel-preview"></a>Criar funções personalizadas no Excel (Visualização)

Funções personalizadas permitem que os desenvolvedores adicionem novas funções para o Excel, definindo essas funções em JavaScript como parte de um suplemento. Os usuários podem então acessar funções personalizadas como qualquer outra função nativa do Excel (como `SUM()`). Este artigo explica como criar as funções personalizadas no Excel.

A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel. A `CONTOSO.ADD42` função personalizada foi projetada para adicionar 42 ao par de números que o usuário especifica como parâmetros de entrada para a função.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

O código a seguir define a função personalizada `ADD42`.

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

As funções personalizadas agora estão disponíveis no Developer Preview para Windows, Mac e Excel Online. Para testá-las, conclua estas etapas:

1. Instale o Office (compilação 10827 no Windows ou 13.329 no Mac) e participe do programa [Office Insider](https://products.office.com/office-insider). Você deve se associar ao programa Office Insider para ter acesso às funções personalizadas; atualmente as funções personalizadas estão desabilitadas em todas as versões do Office, a menos que você seja um membro do programa Office Insider.

2. Use [Yo Office](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento de funções personalizadas do Excel e siga as instruções no [README OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) para usar o projeto.

3. Digite `=CONTOSO.ADD42(1,2)` em qualquer célula de uma planilha do Excel e pressione **Inserir** para executar a função personalizada.

> [!NOTE]
> A seção de [Problemas conhecidos](#known-issues) neste artigo especifica as limitações atuais de funções personalizadas.

## <a name="learn-the-basics"></a>Noções básicas

No projeto de funções personalizadas que você criou usando o [Office Yo](https://github.com/OfficeDev/generator-office), você verá os seguintes arquivos:

| Arquivo | Formato do arquivo | Descrição |
|------|-------------|-------------|
| **./src/customfunctions.js** | JavaScript | Contém o código que define as funções personalizadas. |
| **./config/customfunctions.json** | JSON | Contém metadados que descrevem as funções personalizadas e permitem ao Excel registrar as funções personalizadas para torná-las disponíveis aos usuários finais. |
| **./index.html** | HTML | Fornece uma referência de &lt;script&gt; ao arquivo JavaScript que define as funções personalizadas. |
| **./manifest.xml** | XML | Especifica o namespace para todas as funções personalizadas dentro do suplemento e o local dos arquivos JavaScript, JSON e HTML que estão listados anteriormente nesta tabela. |

### <a name="manifest-file-manifestxml"></a>Arquivo de manifesto (./manifest.xml)

O arquivo de manifesto XML para um suplemento que define as funções personalizadas especifica o namespace para todas as funções personalizadas dentro do suplemento e do local dos arquivos JavaScript, JSON e HTML. A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir em manifesto de um suplemento para habilitar o Excel a executar funções personalizadas.  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> As funções do Excel estão anexadas pelo namespace especificado em seu arquivo de manifesto XML. Um namespace de uma função vem antes do nome da função e são separados por um período. Por exemplo, para chamar a função `ADD42()` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque CONTOSO é o namespace e `ADD42` é o nome da função especificada no arquivo JSON. O prefixo deve ser usado como identificador para a sua empresa ou para o suplemento. 

### <a name="json-file-configcustomfunctionsjson"></a>Arquivo JSON (./config/customfunctions.json)

Um arquivo de metadados de funções personalizadas fornece as informações que o Excel precisa para registrar as funções personalizadas e torná-las disponíveis aos usuários finais. As funções personalizadas são registradas quando um usuário executa o suplemento pela primeira vez. Depois disso, elas estarão disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento inicialmente executou.)

> [!TIP]
> As configurações do servidor para o arquivo JSON devem ter [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para que funções personalizadas funcionem corretamente no Excel Online.

O código a seguir em **customfunctions.json** especifica os metadados para a função `ADD42` descrita anteriormente neste artigo. Esses metadados definem o nome da função, descrição, valor de retorno, parâmetros de entrada e demais dados. A tabela que segue o exemplo de código a seguir fornece informações detalhadas sobre as propriedades individuais dentro desse objeto JSON.

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
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
        }
    ]
}
```

A tabela a seguir lista as propriedades que estão normalmente presentes no arquivo de metadados JSON. Para obter informações mais detalhadas sobre o arquivo de metadados JSON, incluindo opções não usadas no exemplo anterior, consulte [Metadados de funções personalizadas](custom-functions-json.md).

| Propriedade  | Descrição |
|---------|---------|
| `id` | Uma ID exclusiva para a função. Essa ID não deve ser alterada depois de ser definida. |
| `name` | O nome da função que é mostrado no menu de preenchimento automático à medida que um usuário digita uma fórmula dentro de uma célula. No menu de preenchimento automático, esse valor será prefixado pelo namespace das funções personalizadas especificado no arquivo de manifesto XML. |
| `helpUrl` | URL de uma página que é exibida quando o usuário solicita ajuda. |
| `description` | Descreve o que significa a função. Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu Preenchimento Automático dentro do Excel. |
| `result`  | Objeto que define o tipo de informação que é retornado pela função. O valor da propriedade filho `type` pode ser uma **sequência de caracteres**, **número** ou **booleano**. O valor da propriedade filho `dimensionality` pode ser **escalar** ou **matriz** (uma matriz bidimensional de valores do `type` especificado). |
| `parameters` | Matriz que define os parâmetros de entrada para a função. As propriedades filho `name` e `description` são usadas no Intellisense do Excel. As propriedades filho `type` e `dimensionality` são idênticas às propriedades filho do objeto `result` descrito anteriormente nesta tabela. |
| `options` | Permite que você personalize alguns aspectos de como e quando o Excel executa a função. Para obter mais informações sobre como essa propriedade pode ser usada, consulte [Funções em fluxo contínuo](#streamed-functions) e [Cancelamento](#canceling-a-function) mais adiante neste artigo. |

## <a name="functions-that-return-data-from-external-sources"></a>Funções que retornam dados de fontes externas

Se uma função personalizada recupera dados de uma fonte externa, como web, ela deve:

1. Retornar um Promise do JavaScript para o Excel.

2. Resolver o Promise com o valor final usando a função de retorno de chamada.

Funções personalizadas exibem um resultado temporário `#GETTING_DATA` na célula enquanto o Excel aguarda o resultado final. Os usuários podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.

No exemplo de código a seguir, a função personalizada `getTemperature()` recupera a temperatura atual de um termômetro. Observe que `sendWebRequest` é uma função hipotética (não especificada aqui) que usa o XHR para chamar um serviço da Web de temperatura.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>Funções em fluxo contínuo

Funções personalizadas em fluxo contínuo permitem que você transmita para células repetidamente ao longo do tempo, sem exigir que um usuário solicite explicitamente o recálculo. O exemplo de código a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo. Observe o seguinte sobre este código:

- O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.

- O parâmetro final, `handler`, nunca é especificado no seu código de registro e nunca é exibido no menu de preenchimento automático para usuários do Excel ao inserir a função. É a função de retorno de chamada `setResult` usada para passar dados da função para o Excel para atualizar o valor de uma célula.

- Para que o Excel passe a função `setResult` no objeto `handler`, você deve declarar suporte para fluxo contínuo durante o registro da função, definindo a opção `"stream": true` na propriedade `options` para a função personalizada no arquivo de metadados JSON.

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a>Cancelamento de uma função

Em alguns casos, talvez seja necessário cancelar a execução de uma função personalizada em fluxo contínuo para reduzir seu consumo de largura de banda, a memória de trabalho e a carga da CPU. O Excel cancela a execução de uma função nas seguintes situações:

- Quando o usuário edita ou exclui uma célula que faz referência à função.

- Quando um dos argumentos (entradas) para a função é alterado. Nesse caso, uma nova chamada de função é disparada após o cancelamento.

- O usuário aciona manualmente um recálculo. Nesse caso, uma nova chamada de função é disparada após o cancelamento.

> [!NOTE]
> Você deve implementar um manipulador de cancelamento para todas as funções de fluxo contínuo.

Para tornar uma função cancelável, defina a opção `"cancelable": true` na propriedade `options` para a função personalizada no arquivo de metadados JSON.

O código a seguir mostra a mesma função `incrementValue` descrita anteriormente, mas desta vez com um manipulador de cancelamento implementado. Neste exemplo, `clearInterval()` será executado quando a função `incrementValue` for cancelada.

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a>Compartilhamento e salvamento de estado

Funções personalizadas podem salvar os dados em variáveis JavaScript globais. Em chamadas subsequentes, sua função personalizada pode usar valores salvos nessas variáveis. O estado salvo é útil quando os usuários adicionam a mesma função personalizada a mais de uma célula, porque todas as instâncias da função podem compartilhar o estado. Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.

O código a seguir mostra uma implementação da função anterior de fluxo contínuo de temperatura que salva o estado de forma global. Observe o seguinte sobre este código:

- `refreshTemperature` é uma função de fluxo contínuo que lê a temperatura de um determinado termômetro a cada segundo. Novas temperaturas são salvas na variável `savedTemperatures`, mas o valor da célula não é atualizado diretamente. Não deve ser chamada diretamente de uma célula da planilha, *por isso não está registrada no arquivo JSON*.

- `streamTemperature` atualiza os valores de temperatura exibidos na célula a cada segundo e usa variável `savedTemperatures` como fonte de dados. Deve ser registrada no arquivo JSON e nomeada com todas as letras maiúsculas, `STREAMTEMPERATURE`.

- Os usuários podem chamar `streamTemperature` de várias células na interface de usuário do Excel. Cada chamada lê dados da mesma variável `savedTemperatures`.

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
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

## <a name="working-with-ranges-of-data"></a>Trabalhar com intervalos de dados

Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou ela pode retornar um intervalo de dados. No JavaScript, um intervalo de dados é representado como uma matriz bidimensional.

Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel. A função a seguir aceita o parâmetro `values`, que é do tipo `Excel.CustomFunctionDimensionality.matrix`. Note que no metadados JSON para esta função, você definiria a propriedade do parâmetro `type` para `matrix`.

```js
function secondHighest(values){
     let highest = values[0][0], secondHighest = values[0][0];
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

## <a name="handling-errors"></a>Lidar com erros

Quando você criar um suplemento que define funções personalizadas, certifique-se de incluir a lógica de manipulação de erros para considerar os erros de tempo de execução. A manipulação de erros de funções personalizadas é a mesma que a manipulação de erros [para a API do JavaScript Excel em geral](excel-add-ins-error-handling.md). No exemplo de código a seguir, `.catch` manipulará os erros que ocorreram anteriormente no código.

```js
function getComment(x) {
    //this delivers a section of lorem ipsum from the jsonplaceholder API
    let url = "https://jsonplaceholder.typicode.com/comments/" + x;

    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="known-issues"></a>Problemas conhecidos

- As URLs de ajuda e descrições de parâmetros ainda não são usadas pelo Excel.
- Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.
- Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não são suportadas.
- A implantação por meio do Portal de Administração do Office 365 e do AppSource ainda não está habilitada.
- Funções personalizadas no Excel Online podem parar de funcionar durante uma sessão após um período de inatividade. Atualize a página do navegador (F5) e insira novamente uma função personalizada para restaurar o recurso.
- Se você tiver vários suplementos em execução no Excel para Windows, você poderá ver o resultado temporário **#GETTING_DATA** dentro de células de uma planilha. Feche todas as janelas do Excel e reinicie o Excel.
- Outras ferramentas de depuração para funções personalizadas podem estar disponíveis no futuro. Enquanto isso, você pode depurar no Excel Online usando as ferramentas de desenvolvedor F12. Consulte mais detalhes em [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md).

## <a name="changelog"></a>Log de mudanças

- **7 de novembro de 2017**: Enviados* exemplos e versão prévia de funções personalizadas
- **20 de Nov de 2017**: correção de bug de compatibilidade para quem usa as versões 8801 e posteriores
- **28 de novembro de 2017**: Enviado* suporte para cancelamento em funções assíncronas (requer alteração para funções de fluxo contínuo)
- **7 de maio de 2018**: Enviado* o suporte para Mac, Excel Online e funções síncronas executadas dentro processo
- **20 de setembro de 2018**: Enviado suporte para tempo de execução do JavaScript para funções personalizadas. Para obter mais informações, consulte [Tempo de execução para funções personalizadas do Excel](custom-functions-runtime.md).

\* para o Canal Office Insiders

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md)