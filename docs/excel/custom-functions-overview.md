---
title: Criar funções personalizadas no Excel (Versão Prévia)
description: ''
ms.date: 01/23/2018
---

# <a name="create-custom-functions-in-excel-preview"></a>Criar funções personalizadas no Excel (Versão Prévia)

Funções personalizadas (semelhantes a funções definidas pelo usuário ou UDFs) permitem que os desenvolvedores adicionem qualquer função JavaScript no Excel usando um suplemento. Os usuários podem acessar as funções personalizadas como qualquer outra função nativa do Excel (como =SOMA()). Este artigo explica como criar as funções personalizadas no Excel.

Veja a aparência das funções personalizadas no Excel:

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

Aqui está o código de um exemplo de função personalizada que soma 42 a um par de números.

```js
function add42 (a, b) {
    return a + b + 42;
}
```

As funções personalizadas agora estão disponíveis na visualização. Siga estas etapas para experimentá-las:

1.  Participe do programa [Office Insider](https://products.office.com/en-us/office-insider) para instalar a versão do Excel 2016 necessária para personalizar funções no computador (versão 16.8711 ou posterior). Você deve escolher canal "Insider" para a visualização de funções personalizadas funcionarem.
2.  Clone o repositório [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) e siga as instruções em *README.md* para iniciar o suplemento no Excel.
3.  Digite `=CONTOSO.ADD42(1,2)` em qualquer célula e pressione **Enter** para executar a função personalizada.
4.  Se tiver dúvidas, faça perguntas no Stack Overflow com a marcação [office-js](https://stackoverflow.com/questions/tagged/office-js).

Confira a seção Problemas conhecidos no final deste artigo, que inclui as limitações atuais das funções personalizadas e que será atualizado ao longo do tempo.

## <a name="learn-the-basics"></a>Noções básicas


No repositório de exemplo clonado, você verá os seguintes arquivos:

-   *customfunctions.js*, que contém:

    -   O código da função personalizada a ser adicionada no Excel.
    -   O código de registro para conectar sua função personalizada ao Excel. O registro faz com que suas funções personalizadas apareçam na lista de funções disponíveis exibida quando os usuários digitam nas células.
-   *customfunctions.html*, que fornece uma referência de &lt;Script&gt; ao *customfunctions.js*. Este arquivo não é exibido na interface do usuário do Excel.
-   *manifest.xml*, que informa ao Excel a localização dos arquivos HTML e JS necessários para executar as funções personalizadas.

### <a name="javascript-file-customfunctionsjs"></a>Arquivo JavaScript (*customfunctions.js*)

O código a seguir em customfunctions.js declara a função personalizada `add42` e registra a função no Excel.

```js
function add42 (a, b) {
    return a + b + 42;
}

Excel.Script.customFunctions["CONTOSO"]["ADD42"] = {
    call: add42,
    description: "Adds 42 to the sum of two numbers",
    helpUrl: "https://www.contoso.com/help.html",
    result: {
        resultType: Excel.CustomFunctionValueType.number,
        resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    parameters: [{
        name: "num 1",
        description: "The first number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    {
        name: "num 2",
        description: "The second number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    }],
    options:{ batch: false, stream: false }
};

Excel.run(function(ctx) {
    ctx.workbook.customFunctions.addAll();
});
```

O **registro** da função personalizada usa o bloco de código `Excel.Script.customFunctions["CONTOSO"]["ADD42"]`. Você precisa dos seguintes parâmetros para registrar a função no Excel:

-   Prefixo e nome da função: O primeiro valor em `Excel.Script.customFunctions` é o prefixo (nesse caso, CONTOSO é o prefixo). O segundo valor no `Excel.Script.customFunctions` é o nome de função (nesse caso, ADD42 é o nome da função). No Excel, o prefixo e o nome da função são separados por um ponto: para usar a função personalizada, combine o prefixo da função (CONTOSO) com o nome da função (ADD42) e insira `=CONTOSO.ADD42` em uma célula. Por convenção, prefixos e nomes de funções usam letras maiúsculas. O prefixo deve ser usado como identificador para o suplemento.
-   `call`: Define a função JavaScript a ser chamada (por exemplo, `add42`). O nome da função JavaScript não precisa corresponder ao nome registrado no Excel.
-   `description`: A descrição aparece no menu de preenchimento automático do Excel.
-   `helpUrl`: Quando o usuário solicitar a ajuda para uma função, o Excel abre um painel de tarefas e exibe a página da Web encontrada neste URL.
-   `result`: Define o tipo de informação retornada pela função do Excel.

    -   `resultType`: Sua função pode retornar uma `"string"` ou um `"number"` (também usados para datas e moedas). Confira mais informações em [Enumerações das funções personalizadas](https://dev.office.com/reference/add-ins/excel/customfunctionsenumerations).
    -   `resultDimensionality`: Sua função poderá retornar um único valor (`"scalar"`) ou uma `"matrix"` de valores. Ao retornar uma matriz de valores, a função retornará uma matriz, onde cada elemento da matriz é outra matriz que representa uma linha de valores. Confira mais informações em [Enumerações das funções personalizadas](https://dev.office.com/reference/add-ins/excel/customfunctionsenumerations). O exemplo a seguir retorna uma matriz de 2 colunas e 3 linhas de valores de uma função personalizada.

        ```js
        return [["first","row"],["second","row"],["third","row"]];
        ```

-   Sua função personalizada pode conter argumentos como entrada. Os argumentos passados para sua função personalizada são especificados na propriedade *parameters*. A ordem dos parâmetros na definição deve corresponder à ordem na função JavaScript. Para cada parâmetro, defina as seguintes propriedades:

    -   `name`: A cadeia de caracteres exibida no Excel para representar o parâmetro.
    -   `description`: A cadeia de caracteres exibida para fornecer mais informações sobre o parâmetro.
    -   `valueType`: Um `"number"` ou uma `"string"`, semelhante à propriedade resultType descrita acima.
    -   `valueDimensionality`: Um valor `"scalar"` ou uma `"matrix"` de valores, semelhante à propriedade resultDimensionality descrita anteriormente. Parâmetros do tipo matriz permitem que o usuário selecione intervalos maiores do que uma única célula.

-   `options`: permite tipos especiais de funções personalizadas descritas com mais detalhes neste artigo.

Para realizar um registro completo de todas as funções definidas usando `Excel.Script.customFunctions`, certifique-se de chamar `CustomFunctions.addAll()`.

Após o registro, funções personalizadas serão disponibilizadas em todas as pastas de trabalho (não apenas aquela onde o suplemento foi iniciado inicialmente) para um usuário. As funções são exibidas no menu de preenchimento automático quando o usuário começa a digitá-la. Durante o desenvolvimento e o teste, você pode limpar manualmente o cache do seu computador dos metadados de registro ao excluir a pasta `<user>\AppData\Local\Microsoft\Office\16.0\Wef\CustomFunctions`.


### <a name="manifest-file-manifestxml"></a>Arquivo de manifesto (*manifest.xml*)

O exemplo a seguir em manifest.xml permite que o Excel localize o código de suas funções.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">

    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="scriptURL" />
                        <!— Required. The Developer Preview does not use the Script element.-->
                    </Script>
                    <Page>
                        <SourceLocation resid="pageURL"/>
                    </Page>
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>

    <Resources>
        <bt:Urls>
            <bt:Url id="scriptURL" DefaultValue="https://www.contoso.com/addin/customfunctions.js" />
            <bt:Url id="pageURL" DefaultValue="https://www.contoso.com/addin/customfunctions.html" />
        </bt:Urls>
    </Resources>

</VersionOverrides>

```

O código anterior especifica:

-   Um elemento `<Script>`, que é obrigatório, mas não é usado na versão Developer Preview.
-   Um elemento `<Page>` vinculado à página HTML do suplemento. A página HTML inclui uma referência de &lt;Script&gt; para o arquivo JavaScript (*customfunctions.js*) que contém a função personalizada e o código de registro. A página HTML é uma página oculta e nunca é exibida na interface de usuário.

## <a name="asynchronous-functions"></a>Funções assíncronas

Se sua função personalizada recuperar dados da web, será necessário fazer uma chamada assíncrona para obtê-los. Ao chamar serviços Web externos, a função personalizada deve:

1.   Retornar uma Promise do JavaScript para o Excel.
2.   Realizar a solicitação HTTP para chamar o serviço externo.
3.   Resolver a promessa usando o retorno de chamada `setResult`. `setResult` envia o valor para o Excel.

O código a seguir mostra um exemplo de uma função personalizada que recupera a temperatura de um termômetro.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult, setError){
        sendWebRequestExample(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>Funções de streaming

Funções personalizadas de streaming permitem que você insira dados em células repetidamente ao longo do tempo, sem precisar esperar que o Excel ou os usuários solicitem recálculos. Por exemplo, a função personalizada `incrementValue` no código a seguir adiciona um número ao resultado a cada segundo e o Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`. Para ver o código de registro usado com `incrementValue`, leia o arquivo *customfunctions.js*.

```js
function incrementValue(increment, caller){ 
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

Para funções de streaming, o parâmetro final, `caller`, não é especificado no código de registro e nunca é exibido no menu de preenchimento automático para usuários do Excel ao inserir a função. É a função de retorno de chamada `setResult` usada para passar dados da função para o Excel para atualizar o valor de uma célula. Para que o Excel passe a função `setResult` no objeto `caller`, é necessário declarar o suporte para streaming durante o registro da função ao configurar o parâmetro `stream` para `true`.

## <a name="cancellation"></a>Cancelamento

Você pode cancelar funções e funções assíncronas de streaming. É importante cancelar as chamadas de função para reduzir o consumo de largura de banda, a memória de trabalho e a carga da CPU. O Excel cancela chamadas de funções nas seguintes situações:
- O usuário edita ou exclui uma célula que faz referência à função.
- É alterado um dos argumentos (entradas) para a função. Nesse caso, uma nova chamada de função é disparada, além do cancelamento.
- O usuário aciona manualmente um recálculo. Como no caso acima, uma nova chamada de função é disparada, além do cancelamento.

O código a seguir mostra o exemplo anterior com o cancelamento implementado. No código, o objeto `caller` contém uma função `onCanceled` que deve ser definida para cada função personalizada.

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

## <a name="saving-state"></a>Salvar estado

Funções personalizadas podem salvar os dados em variáveis JavaScript globais. Em chamadas subsequentes, sua função personalizada pode usar valores salvos nessas variáveis. O estado salvo é útil quando os usuários inserem várias instâncias da mesma função personalizada e precisam compartilhar dados entre si. Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.

O código a seguir mostra uma implementação da função de streaming de temperatura anterior que salva o estado usando a variável `savedTemperatures`. O código demonstra os conceitos a seguir:

-   **Salvar dados.** `refreshTemperature` é uma função de streaming que lê a temperatura de um determinado termômetro a cada segundo. Novas temperaturas são salvas na variável savedTemperatures.

-   **Usar dados salvos.** `streamTemperature` atualiza os valores de temperatura exibidos na interface de usuário do Excel a cada segundo. Temperaturas são lidas em `savedTemperature` e enviadas para a interface de usuário do Excel usando `setResult`. Os usuários podem chamar `streamTemperature` de várias células na interface de usuário do Excel. Cada chamada para `streamTemperature` lerá os dados de `savedTemperatures`.

> Nesse caso, nós registramos `streamTemperature` como a função personalizada no Excel.

```js
var savedTemperatures{};

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID);
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequestExample(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a>Trabalhar com intervalos de dados

Sua função personalizada pode levar a um intervalo de dados como um parâmetro ou você pode retornar um intervalo de dados de uma função personalizada.

Por exemplo, suponha que sua função retorne a segunda maior temperatura de um intervalo de valores de temperatura armazenados no Excel. A função a seguir usa o parâmetro `temperatures`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.

```js
function secondHighestTemp(temperatures){ 
     var highest = -273, secondHighest = -273;
     for(var i = 0; i < temperatures.length; i++){
         for(var j = 0; j < temperatures[i].length; j++){
             if(temperatures[i][j] <= highest){
                 secondHighest = highest;
                 highest = temperatures[i][j];
             }
             else if(temperatures[i][j] <= secondHighest){
                 secondHighest = temperatures[i][j];
             }
         }
     }
     return secondHighest;
 }
```

Se você cria uma função que retorna um intervalo de dados, é necessário inserir uma fórmula de matriz no Excel para ver todo o intervalo de valores. Para saber mais, confira [Diretrizes e exemplos de fórmulas de matrizes](https://support.office.com/pt-br/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7).

## <a name="known-issues"></a>Problemas conhecidos

Os recursos a seguir ainda não têm suporte na versão Developer Preview.

-   Processamento em lotes, que permite a agregação de várias chamadas na mesma função para melhorar o desempenho.

-   As descrições de URLs e parâmetros de Ajuda ainda não são usadas pelo Excel.

-   A publicação de suplementos que usam funções personalizadas no AppSource ou uma implantação centralizada do Office 365.

-   Funções personalizadas não estão disponíveis no Excel para Mac, no Excel para iOS e no Excel Online.

-   Atualmente, os suplementos dependem de um processo de navegador oculto para executar funções personalizadas. No futuro, o JavaScript será executado diretamente em algumas plataformas para garantir que as funções personalizadas sejam mais rápidas e usem menos memória. Além disso, a página HTML referenciada pelo elemento &lt;Página&gt; no manifesto não será necessária para a maioria das plataformas, já que o Excel executa o JavaScript diretamente. Para se preparar para essa alteração, certifique-se de que suas funções personalizadas não usem o DOM da página da Web.

## <a name="changelog"></a>Log de mudanças

- **7 de novembro de 2017**: enviados exemplos e visualizações de funções personalizadas
- **20 de Nov de 2017**: correção de bug de compatibilidade para quem usa as versões 8801 e posteriores
- **28 de novembro de 2017**: enviado o suporte para cancelamento em funções assíncronas (requer a alteração de funções de streaming)
