---
ms.date: 09/27/2018
description: Criar uma função personalizada no Excel usando o JavaScript.
title: Criar funções personalizadas no Excel (visualização)
ms.openlocfilehash: f6b658bbd119a785b342ec22bc1b341f6902da3f
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459340"
---
# <a name="create-custom-functions-in-excel-preview"></a>Criar funções personalizadas no Excel (versão prévia)

Funções personalizadas permitem que os desenvolvedores adicionem novas funções ao Excel, definindo-as em JavaScript como parte de um suplemento. Os usuários no Excel podem acessar funções personalizadas tal como fariam com qualquer função nativa do Excel, como `SUM()`. Este artigo explica como criar funções personalizadas no Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel. A função personalizada `CONTOSO.ADD42` foi projetada para adicionar 42 ao par de números especificado pelo usuário como os parâmetros de entrada para a função.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

O código a seguir define a função personalizada `ADD42`.

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> A seção de [Problemas conhecidos](#known-issues) mais adiante neste artigo especifica as limitações atuais das funções personalizadas.

## <a name="components-of-a-custom-functions-add-in-project"></a>Componentes de um projeto de suplemento de funções personalizadas

Se você usar o [gerador de Yo Office](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento de funções personalizadas do Excel, verá os seguintes arquivos no projeto criado pelo gerador:

| Arquivo | Formato do arquivo | Descrição |
|------|-------------|-------------|
| **./src/customfunctions.js**<br/>ou<br/>**./src/customfunctions.ts** | JavaScript<br/>ou<br/>TypeScript | Contém o código que define as funções personalizadas. |
| **./config/customfunctions.json** | JSON | Contém metadados que descrevem as funções personalizadas e permitem que o Excel registre as funções personalizadas e as disponibilize para o usuário final. |
| **./index.html** | HTML | Fornece uma referência de &lt;script&gt; ao arquivo JavaScript que define as funções personalizadas. |
| **./manifest.xml** | XML | Especifica o namespace para todas as funções personalizadas dentro do suplemento e o local dos arquivos JavaScript, JSON e HTML listados anteriormente nesta tabela. |

As seções a seguir fornecem mais informações sobre esses arquivos.

### <a name="script-file"></a>Arquivo de script 

O arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** no projeto criado pelo gerador Yo Office) contém o código que define as funções personalizadas e mapeia os nomes das funções personalizadas para os objetos no [arquivo de metadados JSON](#json-metadata-file). 

Por exemplo, o código a seguir define as funções personalizadas `add` e `increment` e, em seguida, especifica informações de mapeamento para ambas as funções. A função `add` é mapeada para o objeto no arquivo de metadados JSON em que o valor da propriedade `id` é **ADD**, e a função `increment` é mapeada para o objeto no arquivo de metadados em que o valor da propriedade `id` é **INCREMENT**. Consulte as [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para obter mais informações sobre o mapeamento de nomes de funções no arquivo de script para objetos no arquivo de metadados JSON.

```js
function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

// map `id` values in the JSON metadata file to the JavaScript function names
CustomFunctionMappings.ADD = add;
CustomFunctionMappings.INCREMENT = increment;
```

### <a name="json-metadata-file"></a>Arquivo de metadados JSON 

O arquivo de metadados de funções personalizadas (**./config/customfunctions.json** no projeto criado pelo gerador Yo Office) fornece as informações que o Excel precisa para registrar as funções personalizadas e disponibilizá-las para os usuários finais. Funções personalizadas serão registradas quando um usuário executa um suplemento pela primeira vez. Depois disso, eles ficam disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento foi inicialmente executado.)

> [!TIP]
> As configurações do servidor que hospeda o arquivo JSON devem ter [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para que funções personalizadas funcionem corretamente no Excel Online.

O código a seguir no **customfunctions.json** especifica os metadados para as funções `add` e  `increment`, descritas anteriormente. A tabela apresentada na sequência do exemplo código a seguir fornece informações detalhadas sobre as propriedades individuais dentro desse objeto JSON. Consulte [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para obter mais informações sobre como especificar o valor das propriedades `id` e `name` no arquivo de metadados JSON.

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      "description": "Periodically increment a value",
      "helpUrl": "http://www.contoso.com",
      "result": {
          "type": "number",
          "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "increment",
            "description": "Amount to increment",
            "type": "number",
            "dimensionality": "scalar"
        }
    ],
    "options": {
        "cancelable": true,
        "stream": true
      }
    }
  ]
}
```

A tabela a seguir lista as propriedades normalmente presentes no arquivo de metadados JSON. Para obter informações mais detalhadas sobre o arquivo de metadados JSON, consulte [Metadados de funções personalizadas](custom-functions-json.md).

| Propriedade  | Descrição |
|---------|---------|
| `id` | Uma ID exclusiva para a função. Essa ID não deve ser alterada depois de definida. |
| `name` | O nome da função exibido para os usuários finais no Excel. No Excel, esse nome de função será prefixado pelo namespace das funções personalizadas especificado no [arquivo de manifesto XML](#manifest-file). |
| `helpUrl` | URL da página que é exibida quando o usuário solicita ajuda. |
| `description` | Descreve o que a função faz. Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu de preenchimento automático dentro do Excel. |
| `result`  | Objeto que define o tipo de informação retornada pela função. O valor da propriedade filha `type` pode ser **string**, **number** ou **boolean**. O valor da propriedade filha `dimensionality` pode ser **scalar** ou **matrix** (uma matriz bidimensional de valores do `type` especificado). |
| `parameters` | Uma matriz que define os parâmetros de entrada da função. As propriedades filhas `name` e `description` aparecem no intelliSense do Excel. O valor da propriedade filha `type` pode ser **string**, **number** ou **boolean**. O valor da propriedade filha `dimensionality` pode ser **scalar** ou **matrix** (uma matriz bidimensional de valores do `type` especificado). |
| `options` | Permite personalizar alguns aspectos da como e quando o Excel executa a função. Para obter mais informações sobre como essa propriedade pode ser usada, consulte [Funções de fluxo contínuo](#streaming-functions) e [Cancelar uma função](#canceling-a-function) mais à frente neste artigo. |

### <a name="manifest-file"></a>Arquivo de manifesto

O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto que criado pelo gerador Yo Office) especifica o namespace para todas as funções personalizadas dentro do suplemento e o local dos arquivos JavaScript, JSON e HTML. A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Funções do Excel são pré-inseridas pelo namespace especificado em seu arquivo de manifesto XML. O namespace de uma função vem antes do nome dela e é separado por um ponto. Por exemplo, para chamar a função `ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque a CONTOSO é o namespace e `ADD42` é o nome da função especificado no arquivo JSON. O namespace funciona como um identificador para a sua empresa ou para o suplemento. 

## <a name="functions-that-return-data-from-external-sources"></a>Funções que retornam dados de fontes externas

Se uma função personalizada recupera dados de uma fonte externa, como web, ela deve:

1. Retornar um Promise do JavaScript para o Excel.

2. Resolver a Promessa com o valor final usando a função de retorno de chamada.

Funções personalizadas exibem um resultado temporário `#GETTING_DATA` na célula enquanto o Excel aguarda o resultado final. Os usuários podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.

No exemplo de código a seguir, a função personalizada `getTemperature()` recupera a temperatura atual de um termômetro. Observe que `sendWebRequest` é uma função hipotética (não especificada aqui) que usa [XHR](custom-functions-runtime.md#xhr-example) para chamar um serviço Web de temperatura.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a>Funções de fluxo contínuo

Funções personalizadas de fluxo contínuo permitem múltiplos dados de saída em uma célula ao longo do tempo, sem precisar que o usuário explicitamente solicite a atualização dos dados. O exemplo de código a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo. Observe o seguinte sobre este código:

- O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.

- O segundo parâmetro de entrada, `handler`, não é exibido para o usuário final no Excel quando ele seleciona a função no menu de preenchimento automático.

- O retorno de chamada `onCanceled` define a função que é executada quando a função for cancelada. Você deve implementar um manipulador de cancelamento como este para qualquer função de fluxo contínuo. Para obter mais informações, consulte [Cancelar uma função](#canceling-a-function). 

```js
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}
```

Quando você especifica os metadados de uma função de fluxo contínuo no arquivo de metadados JSON, deve definir as propriedades `"cancelable": true` e `"stream": true` dentro do objeto `options`, conforme mostrado no exemplo a seguir.

```json
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a>Cancelar uma função

Em alguns casos, talvez seja necessário cancelar a execução de uma função personalizada de fluxo contínuo para reduzir seu consumo de largura de banda, a memória de trabalho e a carga da CPU. O Excel cancela a execução de uma função nas seguintes situações:

- Quando o usuário edita ou exclui uma célula que faz referência à função.

- Quando um dos argumentos (entradas) da função é alterado. Nesse caso, uma nova chamada de função é acionada após o cancelamento.

- Quando o usuário aciona o recálculo manualmente. Nesse caso, uma nova chamada de função é acionada após o cancelamento.

Para habilitar a capacidade de cancelar uma função, você deve implementar um manipulador de cancelamento dentro da função do JavaScript e especificar a propriedade `"cancelable": true` dentro do objeto `options` nos metadados JSON que descrevem a função. Os exemplos de código na seção anterior deste artigo fornecem um exemplo dessas técnicas.

## <a name="saving-and-sharing-state"></a>Compartilhar e salvar um estado

Funções personalizadas podem salvar dados em variáveis globais do JavaScript. Em chamadas subsequentes, sua função personalizada pode usar os valores salvos nessas variáveis. Um estado salvo é útil quando usuários adicionam a mesma função personalizada a mais de uma célula, pois todas as instâncias da função podem compartilhar o estado. Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para esse mesmo recurso.

O exemplo de código a seguir mostra a implementação de uma função de fluxo contínuo de temperatura que salva o estado globalmente. Observe o seguinte sobre este código:

- `refreshTemperature` é uma função de fluxo contínuo que lê a temperatura de um determinado termômetro a cada segundo. Novas temperaturas são salvas na variável `savedTemperatures`, mas não atualiza diretamente o valor da célula. Ela não deve ser chamada diretamente de uma célula da planilha, *então não fica registrada no arquivo JSON*.

- `streamTemperature` atualiza os valores de temperatura exibidos na célula a cada segundo e usa a variável `savedTemperatures` como sua fonte de dados. Deve ser registrada no arquivo JSON e nomeada apenas com letras maiúsculas, `STREAMTEMPERATURE`.

- Os usuários podem chamar `streamTemperature` de várias células na interface do usuário do Excel. Cada chamada lê os dados da mesma variável `savedTemperatures`.

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

Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou pode retornar um intervalo de dados. No JavaScript, um intervalo de dados é representado como uma matriz bidimensional.

Por exemplo, suponha que a sua função retorne o segundo valor mais alto de um intervalo de números armazenados no Excel. A função a seguir aceita o parâmetro `values`, que é do tipo `Excel.CustomFunctionDimensionality.matrix`. Observe que nos metadados JSON dessa função, você faria definiria a propriedade `type` do parâmetro como `matrix`.

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

## <a name="handling-errors"></a>Tratamento de erros

Ao construir um suplemento que define funções personalizadas, certifique-se de incluir lógica para tratamento de erros para lidar com erros em tempo de execução. O tratamento de erros para funções personalizadas funciona da mesma forma que [o tratamento de erros para a API JavaScript do Excel de maneira geral](excel-add-ins-error-handling.md). No exemplo de código a seguir, `.catch` tratará quaisquer erros que ocorram anteriormente no código.

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;

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

- As descrições de URLs e parâmetros de Ajuda ainda não são usadas pelo Excel.
- Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.
- Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não são suportadas.
- A implantação por meio do Portal de Administração do Office 365 e do AppSource ainda não está habilitada.
- Funções personalizadas no Excel Online podem parar de funcionar durante uma sessão após um período de inatividade. Atualize a página do navegador (F5) e insira novamente a função personalizada para restaurar o recurso.
- Se você tiver vários suplementos em execução no Excel para Windows, poderá ver o resultado temporário **#GETTING_DATA** nas células de uma planilha. Feche todas as janelas do Excel e reinicie o programa.
- Ferramentas de depuração específicas para funções personalizadas podem ser disponibilizadas futuramente. Enquanto isso, você pode depurar no Excel Online usando as ferramentas de desenvolvedor F12. Veja mais detalhes em [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).

## <a name="changelog"></a>Log de mudanças

- **7 de novembro de 2017**: Enviados* exemplos e versão prévia de funções personalizadas
- **20 de Novembro de 2017**: Correção de bug de compatibilidade para quem usa o build 8801 e posteriores
- **28 de novembro de 2017**: Enviado* suporte para cancelamento em funções assíncronas (requer alteração para funções de fluxo contínuo)
- **7 de maio de 2018**: Enviado* o suporte para Mac, Excel Online e funções síncronas executadas no processo
- **20 de setembro de 2018**: Enviado o suporte para funções personalizadas de tempo de execução do JavaScript. Para obter mais informações, consulte o [Funções personalizadas para tempo de execução do Excel](custom-functions-runtime.md).

\* para o Canal Office Insiders

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)