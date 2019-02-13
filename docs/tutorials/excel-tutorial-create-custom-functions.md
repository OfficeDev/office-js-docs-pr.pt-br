---
title: Tutorial de funções personalizadas do Excel (visualização)
description: Neste tutorial, você criará um suplemento do Excel que contém uma função personalizada que pode executar cálculos e solicitar ou transmitir dados da web.
ms.date: 01/08/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 4ac735e6fc19f13859d07df6cb3d2443e6dfe2fd
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982017"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a>Tutorial: Criar funções personalizadas no Excel (visualização)

Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas como fariam com qualquer função nativa no Excel, como `SUM()`. Você pode criar funções personalizadas que realizam tarefas simples como cálculos ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.

Neste tutorial, você vai:
> [!div class="checklist"]
> * Crie um suplemento de função personalizada usando o [gerador Yeoman de suplementos do Office](https://www.npmjs.com/package/generator-office). 
> * Usar uma função personalizada predefinida para realizar um cálculo simples.
> * Criar uma função personalizada que solicita dados da web.
> * Criar uma função personalizada que transmite os dados da web em tempo real.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>Pré-requisitos

* [Node](https://nodejs.org/en/) (versão 8.0.0 ou posterior)

* [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

* A versão mais recente do [Yeoman](https://yeoman.io/) e do [Yeoman gerador de suplementos do Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Mesmo se você já instalou o gerador Yeoman, recomendamos atualizar seu pacote para a versão mais recente do npm.

* Excel para Windows (versão 1810 64 bits ou posterior) ou o Excel Online

* Ingressar o [programa Office Insider](https://products.office.com/office-insider) (nível**Insider**, anteriormente chamado de "Insider – modo rápido")

## <a name="create-a-custom-functions-project"></a>Criar um projeto com funções personalizadas

 Para começar, você criará o projeto de código para criar o suplemento função personalizada. Os [gerador Yeoman de suplementos do Office](https://www.npmjs.com/package/generator-office) configurará o seu projeto com algumas funções personalizados iniciais que você pode experimentar.

1. Execute o comando a seguir e responda aos prompts da seguinte forma.
    
    ```
    yo office
    ```
    
    * Escolha o tipo de projeto:`Excel Custom Functions Add-in project (...)`
    * Escolha um tipo de script: `JavaScript`
    * Qual será o nome do suplemento? `stock-ticker`
    
    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/12-10-fork-cf-pic.jpg)
    
    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node.js de suporte.

2. Vá até a pasta do projeto.
    
    ```
    cd stock-ticker
    ```

3. Confie no certificado autoassinado necessário para executar este projeto. Para obter instruções detalhadas para Windows ou Mac, confira [Adicionando Certificados Autoassinados como Certificado Raiz Confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Crie um projeto.
    
    ```
    npm run build
    ```

5. Inicie o servidor local da web, que é executado no Node.js. Você pode experimentar o suplemento função personalizada no Excel para Windows ou o Excel Online.

# <a name="excel-for-windowstabexcel-windows"></a>[Excel para Windows](#tab/excel-windows)

Execute o seguinte comando.

```
npm run start
```

Esse comando inicia o servidor web e sideloads seu suplemento da função personalizada no Excel para Windows.

> [!NOTE]
> Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente. Também é possível habilitar o **[log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** solucionar problemas com arquivo de manifesto XML do add-in, bem como qualquer problema de instalação ou tempo de execução. Gravações de log de tempo de execução `console.log` declarações para um arquivo de log para ajudá-lo a descobrir e corrigir problemas.

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

Execute o seguinte comando.

```
npm run start-web
```

Esse comando inicia o servidor web. Faça o seguinte para sideload o seu suplemento.

<ol type="a">
   <li>No Excel Online, escolha a guia <strong>inserir</strong> pressione e, em seguida, escolha <strong>suplementos</strong>.<br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li>Escolha <strong>Gerenciar Meus suplementos</strong> e selecione <strong>Carregar o Suplemento</strong>.</li> 
   <li>Escolha <strong>Procurar... </strong> e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</li> 
   <li>Selecione o arquivo <strong>manifest. XML</strong> e escolha <strong>abrir</strong>, escolha <strong>Carregar</strong>.</li>
</ol>

> [!NOTE]
> Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente.

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>Experimente uma função personalizada predefinida

O projeto de funções personalizados criados alrady tem duas funções personalizadas predefinidas chamadas INCREMENTO e ADICIONAR. O código para essas funções predefinidas está no arquivo **src/customfunctions.js**. O arquivo **./manifest.xml** especifica que todas as funções personalizadas pertencem a `CONTOSO` namespace. Você usará o namespace CONTOSO para acessar as funções personalizadas no Excel.

Em seguida você vai experimentar a função personalizada `ADD` preenchendo as seguintes etapas:

1. No Excel, vá para qualquer célula e digite `=CONTOSO`. Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.

2. Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.

As `ADD` função personalizada calcula a soma dos dois números que você forneceu e retorna o resultado da **210**.

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Criar uma função personalizada que solicita dados da web

Integração de dados da Web é uma ótima maneira de ampliar o Excel por meio de funções personalizadas. Em seguida, você criará uma função personalizada chamada `stockPrice` que recebe uma citação ações de uma Web API e retorna o resultado para a célula de uma planilha. Esta função personalizada usa IEX Trading API, que é gratuito e não requer autenticação.

1. No projeto**cotações** localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.

2. Em **customfunctions.js**, localize a função `increment` e adicione o seguinte código imediatamente após essa função.

    ```js
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }

> [!NOTE]
> In the January Insiders 1901 Build, there is a bug preventing fetch calls from executing which will result in #VALUE!.
> To workaround this please use the [XMLHTTPRequest API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime#requesting-external-data) to make the web request.

3. In **customfunctions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    O `CustomFunctions.associate` código associa a `id` da função com o endereço de função da `increment` em JavaScript para que o Excel possa ligar para a função.

    Antes que o Excel possa usar a função personalizada, você precisa descrever usando metadados. Você precisa definir a `id` usada no método `associate` anteriormente, além de outros metadados.


4. Abra o arquivo **config/customfunctions.json**. Adicione o seguinte objeto JSON à matriz 'funções' e salve o arquivo.

    ```JSON
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    Este JSON descreve a função `stockPrice`, seus parâmetros e o tipo de resultado ela retornará.

5. Registre novamente o suplemento no Excel para que a nova função esteja disponível. 

# <a name="excel-for-windowstabexcel-windows"></a>[Excel para Windows](#tab/excel-windows)

1. Feche o Excel e abra novamente o Excel.

2. No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)

3. Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.
    ![Insira a faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Insira a faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**. 

3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman. 

4. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

--- 

<ol start="6">
<li> Agora, vamos experimentar a nova função. Na célula <strong>B1</strong>, digite o texto <strong>= da CONTOSO. STOCKPRICE("msft")</strong> e pressione enter. Você verá que o resultado na célula <strong>B1</strong> é o preço atual das ações para uma ação da Microsoft.</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>Criar uma função personalizada assíncrona de streaming

A `stockPrice` função que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços das ações estão sempre mudando. Em seguida, você criará uma função personalizada chamada `stockPriceStream` esse é o preço de uma ação a cada 1000 milissegundos.

1. No projeto**cotações**, adicione o código a seguir **src/customfunctions.js** e salve o arquivo.

    ```js
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    Antes que o Excel possa usar a função personalizada, você precisa descrever usando metadados.
    
2. No projeto **cotações** adicione o seguinte objeto a `functions` matriz dentro do arquivo **config/customfunctions.json** e salve o arquivo.
    
    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

    Este JSON descreve a função `stockPriceStream`. Para qualquer função streaming, a propriedade `stream` e a propriedade `cancelable` devem ser definidas como `true` dentro do objeto `options`, como mostra este exemplo código.

3. Registre novamente o suplemento no Excel para que a nova função esteja disponível.

# <a name="excel-for-windowstabexcel-windows"></a>[Excel para Windows](#tab/excel-windows)

1. Feche o Excel e abra novamente o Excel.

2. No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)

3. Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.
    ![Insira a faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Insira a faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

4. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

--- 

<ol start="4">
<li>Agora, vamos experimentar a nova função. Na célula <strong>C1</strong>, digite o texto <strong>= da CONTOSO. STOCKPRICESTREAM("msft")</strong> e pressione enter. Desde que o mercado de ações esteja aberto, você verá que o resultado na célula <strong>C1</strong> é constantemente atualizado para refletir o preço em tempo uma ação das ações da Microsoft.</li>
</ol>


## <a name="next-steps"></a>Próximas etapas

Parabéns! Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função predefinida, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que transmite dados em tempo real da Web. Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:

> [!div class="nextstepaction"]
> [Criar funções personalizadas no Excel](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>Informações legais

Dados gratuito fornecidos pela [IEX](https://iextrading.com/developer/). Modo de exibição [termos de uso IEX](https://iextrading.com/api-exhibit-a/). O uso da Microsoft dA API IEX neste tutorial é apenas para fins educacionais.


