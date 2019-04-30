---
ms.date: 04/20/2019
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel (versão prévia)
localization_priority: Priority
ms.openlocfilehash: 634b76ed90a30c7aa8252da346ba3f95684967a4
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353248"
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

Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar funções personalizadas em um projeto do Excel, você encontrará que cria os arquivos que controlam as funções, o painel de tarefas e o suplemento geral. Vamos nos concentrar em arquivos que são importantes para funções personalizadas: 

| File | Formato de arquivo | Descrição |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>ou<br/>**./src/functions/functions.ts** | JavaScript<br/>ou<br/>TypeScript | Contém o código que define funções personalizadas. |
| **./src/functions/functions.html** | HTML | Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas. |
| **./manifest.xml** | XML | Especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos JavaScript e HTML listados anteriormente nesta tabela. Também lista os locais de outros arquivos, que o suplemento pode fazer uso, como os arquivos do painel de tarefas e arquivos de comando. |

### <a name="script-file"></a>Arquivo de script

O arquivo de script (**./src/functions/functions.js** ou **./src/functions/functions.ts** no projeto que o gerador Yo Office cria) contém o código que define funções personalizadas, comentários que definem a função, e associa os nomes das funções personalizadas a objetos no arquivo de metadados JSON.

Por exemplo, o código a seguir define funções personalizadas `add` e especifica as informações de mapeamento para as duas funções. Para saber mais, confira [práticas recomendadas de funções personalizados](custom-functions-best-practices.md#associating-function-names-with-json-metadata).

O código a seguir também fornece os comentários de código que definem a função. O necessário `@customfunction` comentário é declarado primeiro, para indicar que se trata de uma função personalizada. Além disso, observe que dois parâmetros foram declarados, `first` e `second`, que é seguido por suas `description` propriedades. Por fim, uma `returns` descrição é fornecida. Para obter mais informações sobre quais comentários são necessários para sua função personalizada, confira [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

### <a name="manifest-file"></a>Arquivo de manifesto

O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto gerador que Yo Office cria) especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON. 

A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas. Se estiver usando o gerador Yo Office, seus arquivos de funções personalizadas gerados conterão um arquivo de manifesto mais complexo, que você pode comparar neste [repositório do Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).

> [!NOTE] 
> As URLs especificadas no arquivo de manifesto para as funções personalizadas JavaScript e JSON e arquivos HTML devem estar publicamente acessíveis e ter o mesmo subdomínio.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> Funções do Excel são anexadas ao namespace especificado no seu arquivo de manifesto XML. O namespace da função vem antes do nome da função e são separados por um ponto. Por exemplo, para acionar a função`ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque `CONTOSO` é o namespace e `ADD42` é o nome da função especificada no arquivo JSON. O namespace deve ser usado como identificador para o as sua empresa ou suplemento. Um namespace pode conter apenas caracteres alfanuméricos e períodos.

## <a name="declaring-a-volatile-function"></a>Como declarar uma função volátil

As [funções voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) são funções nas quais o valor muda de momento a momento, mesmo que nenhum dos argumentos da função tenha mudado. Essas funções são recalculadas sempre que o Excel recalcular. Por exemplo, imagine uma célula que chame a função `NOW`. Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.

O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`. Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

As funções personalizadas permitem que você crie suas próprias funções voláteis, que podem ser úteis ao lidar com datas, horas, números aleatórios e modelagem. Por exemplo, as simulações de Monte Carlo exigem a geração de entradas aleatórias para determinar uma solução ideal.

Para declarar uma função volátil, adicione `"volatile": true` no objeto `options` para a função no arquivo JSON de metadados, como mostra o exemplo a seguir. Observe que uma função não pode ser marcada como `"streaming": true` e `"volatile": true`; em casos em que ambas estejam marcadas com `true`, a opção volátil será ignorada.

```json
{
 "id": "TOMORROW",
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

## <a name="coauthoring"></a>Coautoria

O Excel Online e o Excel para Windows com uma assinatura do Office 365 permitem editar documentos em coautoria, e esse recurso funciona com funções personalizadas. Se a pasta de trabalho usa uma função personalizada, seu colega será solicitado a carregar o suplemento da função personalizada. Depois de carregarem o suplemento, a função personalizada compartilhará resultados por meio de coautoria.

Para saber mais sobre coautoria, confira o tópico [Sobre o recurso de coautoria no Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

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

## <a name="determine-which-cell-invoked-your-custom-function"></a>Determinar quais células chamadas de sua função personalizada

Em alguns casos, você precisará obter o endereço da célula invocada na sua função personalizada. Isso pode ser útil para os seguintes tipos de cenários:

- Formatação de intervalos: Use o endereço da célula como a chave para armazenar informações em [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Em seguida, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `AsyncStorage`.
- Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `AsyncStorage` usando `onCalculated`.
- Reconciliação: Use o endereço da célula para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.

As informações sobre o endereço de uma célula serão expostas somente se `requiresAddress` estiver marcado como `true` no arquivo de metadados JSON da função. A seguir, um exemplo disso para se você fosse gravar esse arquivo JSON manualmente. Você também pode usar a tag `@requiresAddress` se gerar automaticamente seu arquivo JSON. Para mais detalhes, confira [Geração automática do JSON](custom-functions-json-autogeneration.md).

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

## <a name="known-issues"></a>Problemas conhecidos

Veja os problemas conhecidos no nosso [GitHub de funções do Excel personalizado repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Log de alteração de funções personalizadas](custom-functions-changelog.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Depuração de funções personalizadas](custom-functions-debugging.md)
