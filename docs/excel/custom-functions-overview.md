---
ms.date: 12/14/2018
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel (Versão Prévia)
ms.openlocfilehash: 87f56f4c697d19296fe1b539e4071c8e79fbed6a
ms.sourcegitcommit: 09f124fac7b2e711e1a8be562a99624627c0699e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/15/2018
ms.locfileid: "27283113"
---
# <a name="create-custom-functions-in-excel-preview"></a>Criar funções personalizadas no Excel (versão prévia)

Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`. Este artigo descreve como criar as funções personalizadas no Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel. A função personalizada `CONTOSO.ADD42` foi projetada para adicionar 42 ao par dos números que o usuário especifica como parâmetros de entrada para a função.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

O código a seguir define a função personalizada `ADD42`.

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> A seção [Problemas conhecidos](#known-issues) neste artigo especifica as atuais limitações de funções personalizadas.

## <a name="components-of-a-custom-functions-add-in-project"></a>Componentes de um projeto de suplemento de funções personalizadas

Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar funções personalizadas do Excel em um projeto, você verá os seguintes arquivos no projeto que o gerador cria:

| File | Formato de arquivo | Descrição |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>ou<br/>**./src/functions/functions.ts** | JavaScript<br/>ou<br/>TypeScript | Contém o código que define funções personalizadas. |
| **./src/functions/functions.json** | JSON | Contém metadados que descrevem funções personalizadas e permitem que o Excel registre funções personalizadas e as disponibilize para os usuários finais. |
| **./src/functions/functions.html** | HTML | Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas. |
| **./manifest.xml** | XML | Especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON listados anteriormente nesta tabela. |

As seções a seguir fornecem mais informações sobre esses arquivos.

### <a name="script-file"></a>Arquivo de script

O arquivo de script (**./src/functions/functions.js** ou **./src/functions/functions.ts** no projeto gerador que o Yo Office cria) contém o código que define funções personalizadas e mapeia os nomes da funções personalizadas aos objetos em [arquivos de metadados JSON](#json-metadata-file). 

Por exemplo, o código a seguir define funções personalizadas `add` e `increment` e especifica as informações de mapeamento para as duas funções. A `add` função mapeada para o objeto no arquivo nos metadados JSON onde o valor da `id` propriedade **Adicionar**e a `increment` função é mapeada para o objeto no arquivo metadados onde o valor da propriedade`id`é **INCREMENTO**. Ver [Práticas recomendadas de funções personalizados](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para saber mais sobre mapeamento de nomes de função no arquivo de script para objetos no arquivo de metadados JSON.

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

O arquivo de metadados de funções personalizadas (**./src/functions/functions.json** no projeto que o gerador do Yo Office cria) fornece informações exigidas pelo Excel para registrar funções personalizadas e disponibilizá-las aos usuários finais. Funções personalizadas são registradas quando um usuário usar um suplemento pela primeira vez. Depois disso, eles estão disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento foi inicialmente executado.)

> [!TIP]
> Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.

O seguinte código em **functions.json** especifica os metadados para a função `add` e a função `increment` descritas anteriormente. A tabela que segue o código fornece informações detalhadas sobre as propriedades individuais nesse objeto JSON. Ver [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para saber mais sobre como especificar o valor das propriedades`id` e `name` no arquivo de metadados JSON.

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
        "stream": true,
        "volatile": false
      }
    }
  ]
}
```

A tabela a seguir lista as propriedades normalmente presentes no arquivo de metadados JSON. Para saber mais sobre o arquivo de metadados JSON, confira [Metadados de funções personalizadas](custom-functions-json.md).

| Propriedade  | Descrição |
|---------|---------|
| `id` | Identificação exclusiva para a função. Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada. |
| `name` | Nome da função que o usuário final vê no Excel. No Excel, o nome de função será prefixado pelo namespace de funções personalizadas especificado no [arquivo de manifesto XML](#manifest-file). |
| `helpUrl` | A URL da página é exibida quando um usuário solicitar ajuda. |
| `description` | Descreve o que faz a função. Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu de preenchimento automático do Excel. |
| `result`  | Objeto que define o tipo de informação que é retornada pela função do Excel. Para obter informações detalhadas sobre esse objeto, consulte [resultado](custom-functions-json.md#result). |
| `parameters` | Matriz que define os parâmetros de entrada para a função. Para obter informações detalhadas sobre esse objeto, consulte [parâmetros](custom-functions-json.md#parameters). |
| `options` | Permite que você personalize alguns aspectos de como e quando o Excel executa a função. Confira mais informações sobre como essa propriedade pode ser usada em [Função de streaming](#streaming-functions), [Como cancelar uma função](#canceling-a-function) e [Como declarar uma função volátil](#declaring-a-volatile-function) mais adiante neste artigo. |

### <a name="manifest-file"></a>Arquivo de manifesto

O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto gerador que Yo Office cria) especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON. A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
        <Hosts>
            <Host xsi:type="Workbook">
                <AllFormFactors>
                    <ExtensionPoint xsi:type="CustomFunctions">
                        <Script>
                            <SourceLocation resid="Contoso.Functions.Script.Url" />
                        </Script>
                        <Page>
                            <SourceLocation resid="Contoso.Functions.Page.Url"/>
                        </Page>
                        <Metadata>
                            <SourceLocation resid="Contoso.Functions.Metadata.Url" />
                        </Metadata>
                        <Namespace resid="Contoso.Functions.Namespace" />
                    </ExtensionPoint>
                </AllFormFactors>
            </Host>
        </Hosts>
        <Resources>
            <bt:Images>
                <bt:Image id="Contoso.tpicon_16x16" DefaultValue="https://localhost:3000/assets/icon-16.png" />
                <bt:Image id="Contoso.tpicon_32x32" DefaultValue="https://localhost:3000/assets/icon-32.png" />
                <bt:Image id="Contoso.tpicon_80x80" DefaultValue="https://localhost:3000/assets/icon-80.png" />
            </bt:Images>
            <bt:Urls>
                <bt:Url id="Contoso.Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js" />
                <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json" />
                <bt:Url id="Contoso.Functions.Page.Url" DefaultValue="https://localhost:3000/dist/functions.html" />
                <bt:Url id="Contoso.Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
            </bt:Urls>
            <bt:ShortStrings>
                <bt:String id="Contoso.Functions.Namespace" DefaultValue="CONTOSO" />
            </bt:ShortStrings>
        </Resources>
    </VersionOverrides>
```

> [!NOTE]
> Funções do Excel são anexadas ao namespace especificado no seu arquivo de manifesto XML. O namespace da função vem antes do nome da função e são separados por um ponto. Por exemplo, para acionar a função`ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque `CONTOSO` é o namespace e `ADD42` é o nome da função especificada no arquivo JSON. O namespace deve ser usado como identificador para o as sua empresa ou suplemento. Um namespace pode conter apenas caracteres alfanuméricos e períodos.

## <a name="functions-that-return-data-from-external-sources"></a>Funções que retornam os dados de fontes externas

Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:

1. Retornar uma Promise do JavaScript para o Excel.

2. Resolva a promessa com o uso da função retorno de chamada de valor final.

Exibição de funções personalizados mostra um `#GETTING_DATA` resultado temporário na célula enquanto o Excel espera do resultado final. Os usuários podem interagir normalmente com o restante da planilha aguardando o resultado.

No exemplo de código a seguir, a função personalizada `getTemperature()` recupera a temperatura atual de um termômetro. Observe que `sendWebRequest` é uma função hipotética (não especificada aqui) que usa [XHR](custom-functions-runtime.md#xhr-example) para chamar um serviço web de temperatura.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a>Funções Streaming

Funções personalizadas de streaming permitem a saída de dados das células repetidamente ao longo do tempo, sem a necessidade de um usuário explicitamente solicitar a atualização de dados. O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo. Observe o seguinte sobre este código:

- Cada novo valor usando o Excel automaticamente exibirá o `setResult` retorno de chamada.

- O segundo parâmetro de entrada, `handler`, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.

- O `onCanceled` retorno de chamada define a função que é executada quando a função é cancelada. Implemente um identificador de cancelamento assim para qualquer função de streaming. Para saber mais, confira [Cancelar uma função](#canceling-a-function).

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

Quando você especifica os metadados para uma função streaming no arquivo de metadados JSON, você deve definir as propriedades `"cancelable": true` e `"stream": true` no `options` objeto, conforme mostrado no exemplo a seguir.

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

Em algumas situações, talvez seja necessário cancelar a execução de uma função personalizada de streaming para reduzir o consumo de banda larga, memória de trabalho e carregamento de CPU. O Excel cancela a execução de uma função nas seguintes situações:

- Quando o usuário edita ou exclui uma célula que faz referência à função.

- Quando é alterado um dos argumentos (entradas) para a função. Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.

- Quando o usuário aciona manualmente um recálculo. Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.

Para habilitar o recurso cancelar uma função, implemente um identificador de cancelamento dentro da função JavaScript e especifique a propriedade `"cancelable": true` dentro do `options` objeto nos metadados JSON que descreve a função. Amostras de código na seção anterior neste artigo fornecem um exemplo dessas técnicas.

## <a name="declaring-a-volatile-function"></a>Como declarar uma função volátil

As [funções voláteis](https://docs.microsoft.com/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) são funções nas quais o valor muda de momento a momento, mesmo que nenhum dos argumentos da função tenha mudado. Essas funções são recalculadas sempre que o Excel recalcular. Por exemplo, imagine uma célula que chame a função `NOW`. Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.

O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`. Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](https://docs.microsoft.com/pt-BR/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).  
  
As funções personalizadas permitem que você crie suas próprias funções voláteis, que podem ser úteis ao lidar com datas, horas, números aleatórios e modelagem. Por exemplo, as simulações de Monte Carlo exigem a geração de entradas aleatórias para determinar uma solução ideal.  
  
Para declarar uma função volátil, adicione `"volatile": true` no objeto `options` para a função no arquivo JSON de metadados, como mostra o exemplo a seguir. Observe que uma função não pode ser marcada como `"streaming": true` e `"volatile": true`; em casos em que ambas estejam marcadas com `true`, a opção volátil será ignorada.  

```json
{
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## <a name="saving-and-sharing-state"></a>Salvar e compartilhar estado

Funções personalizadas podem salvar os dados em variáveis, que podem ser usadas em chamadas subsequentes. O estado salvo é útil quando os usuários solicitam a mesma função personalizada usando mais de uma célula, porque todas as ocorrências da função podem acessar o estado. Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.

O código a seguir mostra uma implementação da função de streaming de temperatura que salva o estado globalmente. Observe o seguinte sobre este código:

- A função `streamTemperature` atualiza o valor de temperatura exibido na célula a cada segundo e ele usa a variável `savedTemperatures` como fonte de dados.

- Como `streamTemperature` é uma função de streaming, ela implementa um identificador de cancelamento que será executado quando a função for cancelada.

- Se um usuário ligar a função`streamTemperature` de várias células no Excel, a função `streamTemperature` lê os dados a partir da mesma`savedTemperatures` variável toda vez que ela for executada. 

- `refreshTemperature` é uma função de streaming que lê a temperatura de um determinado termômetro a cada segundo e armazena o resultado na variável`savedTemperatures`. Como a função`refreshTemperature` não é exibida para os usuários finais no Excel, não é necessário ser registrado no arquivo JSON.

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
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

Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou pode retornar um intervalo de dados. Em JavaScript, um intervalo de dados é representado como uma matriz bidimensional.

Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel. A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`. Observe que, nos metadados JSON dessa função, você deve definir o parâmetro `type` propriedade para `matrix`.

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

## <a name="discovering-cells-that-invoke-custom-functions"></a>Como descobrir células que invocam funções personalizadas

As funções personalizadas também permitem formatar intervalos, exibir valores armazenados em cache e reconciliar valores usando `caller.address`, o que torna possível descobrir a célula que invocou uma função personalizada. Você pode usar `caller.address` em alguns dos cenários a seguir:

- Formatação de intervalos: use `caller.address` como a chave da célula para armazenar informações em [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Em seguida, use [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `AsyncStorage`.
- Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `AsyncStorage` usando `onCalculated`.
- Reconciliação: use `caller.address` para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.

As informações sobre o endereço de uma célula serão expostas somente se `requiresAddress` estiver marcado como `true` no arquivo de metadados JSON da função. A seguir, um exemplo disso:

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

No arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts**), também será necessário adicionar uma função `getAddress` para encontrar o endereço de uma célula. Essa função pode ter parâmetros, conforme mostrado no exemplo a seguir como `parameter1`. O último parâmetro sempre será `invocationContext`, um objeto com o local da célula que o Excel passa quando `requiresAddress` é marcado como `true` no arquivo de metadados JSON.

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

Por padrão, os valores retornados de uma função `getAddress` seguem o formato abaixo: `SheetName!CellNumber`. Por exemplo, se uma função foi chamada de uma planilha nomeada Despesas na célula B2, o valor retornado seria `Expenses!B2`.

## <a name="handling-errors"></a>Tratamento de erros

Quando você cria um suplemento que define funções personalizadas certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md). No seguinte exemplo de código `.catch` tratará os erros que ocorreram anteriormente no código.

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
- Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não serão aceitas.
- Implantação por meio do Portal de administração do Office 365 e AppSource ainda não estão habilitados.
- Funções personalizadas no Excel Online podem deixar de funcionar durante uma sessão após um período de inatividade. Atualize a página do navegador (F5) e insira novamente uma função personalizada para restaurar o recurso.
- Você pode ver o resultado temporário **# OBTENDO_DADOS** nas células de uma planilha, se você tiver vários suplementos em execução no Excel para Windows. Feche todas as janelas do Excel e reinicie o Excel.
- Ferramentas de depuração especificas para funções personalizadas podem estar disponíveis no futuro. Enquanto isso, você pode depurar no Excel Online usando ferramentas de desenvolvedor F12. Ver mais detalhes em [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).

## <a name="changelog"></a>Log de mudanças

- **7 de novembro de 2017**: enviados exemplos e visualizações de funções personalizadas
- **20 de Nov de 2017**: correção de bug de compatibilidade para quem usa as versões 8801 e posteriores
- **28 de novembro de 2017**: enviado o suporte para cancelamento em funções assíncronas (requer a alteração de funções de streaming)
- **7 de maio de 2018**: Suporte enviado para Mac, Excel Online e funções síncronas em execução no processo
- **20 de setembro de 2018**: Suporte enviado para funções personalizadas de tempo de execução do JavaScript. Para saber mais, confira [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md).
- **20 de outubro de 2018**: Com o [build do Insider de outubro](https://support.office.com/pt-BR/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), funções personalizadas agora exigem o parâmetro "id" na suas [funções personalizadas metadados](custom-functions-json.md) para área de trabalho do Windows e Online. No Mac, esse parâmetro deve ser ignorado.


Em \* canal[Office Insider](https://products.office.com/office-insider), (anteriormente chamado de "Insider – modo rápido")

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
