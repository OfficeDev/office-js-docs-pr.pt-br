# <a name="tutorial-create-custom-functions-in-excel"></a>Tutorial: Criar funções personalizadas no Excel

## <a name="introduction"></a>Introdução

Funções personalizadas permitem que você adicione novas funções ao Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários no Excel podem acessar funções personalizadas tal como fariam com qualquer função nativa no Excel, como `SUM()`. É possível criar funções personalizadas que executem tarefas simples, como cálculos personalizados ou as tarefas mais complexas, como o fluxo contínuo de dados em tempo real da Web a uma planilha.

Neste tutorial, você irá:
> [!div class="checklist"]
> * Criar um projeto de funções personalizadas usando o gerador Yo Office
> * Usar uma função personalizada pré-criada para executar um cálculo simples
> * Criar uma função personalizada que solicita dados da Web
> * Criar uma função personalizada que transmite dados em tempo real da Web

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>Pré-requisitos

* [Node.js e npm](https://nodejs.org/en/)

* [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

* A última versão do [Yeoman](http://yeoman.io/) e o [gerador Yo Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o comando a seguir via prompt de comando:

    ```bash
    npm install -g yo generator-office
    ```

* Excel para Windows (número da versão 10827 ou posterior) ou Excel Online

* [Ingressar no programa Office Insider](https://products.office.com/office-insider)

## <a name="create-a-custom-functions-project"></a>Criar um projeto de funções personalizadas

Este tutorial começa usando o gerador Yo Office para criar os arquivos que você precisa para seu projeto de funções personalizadas.

1. Execute o comando a seguir e responda aos prompts da forma a seguir.

    ```bash
    yo office
    ```

    * Escolha um tipo de projeto: `Excel Custom Functions Add-in project (...)`
    * Escolha um tipo de script: `JavaScript`
    * Qual será o nome do suplemento? `stock-ticker`

    ![O Yo Office busca prompts de funções personalizadas](../images/yo-office-cfs-stock-ticker-3.png)

    Depois de concluir o assistente, o gerador criará os arquivos do projeto e instalará os componentes do nós de suporte.

2. Navegue até a pasta do projeto.

    ```bash
    cd stock-ticker
    ```

3. Inicie o servidor Web local.

    * Se for usar o Excel para Windows para testar suas funções personalizadas, execute o comando a seguir para iniciar o servidor Web local, inicie o Excel e faça o sideload do suplemento:

        ```bash
        npm start
        ```

    * Se for usar o Excel Online para testar suas funções personalizadas, execute o seguinte comando para iniciar o servidor Web local: 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a>Experimentar uma função personalizada pré-criada

O projeto de funções personalizadas que você criou usando o gerador Yo Office contém algumas funções personalizadas pré-criadas, definidas no arquivo **src/customfunction.js**. O arquivo **manifest.xml** no diretório raiz do projeto especifica que todas as funções personalizadas pertencem ao namespace `CONTOSO`.

Antes de poder usar qualquer uma das funções personalizadas pré-criadas, é preciso registrar o suplemento de funções personalizadas no Excel. Faça isso concluindo as etapas para a plataforma que você usará neste tutorial.

* Se for usar o Excel para Windows para testar suas funções personalizadas:

    1. No Excel, escolha a guia **Inserir**, depois escolha a seta para baixo localizada à direita de **Meus suplementos**.  ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)

    2. Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.
        ![Inserir a faixa de opções no Excel para Windows com o suplemento Funções personalizadas do Excel destacado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)

* Se for usar o Excel Online para testar suas funções personalizadas: 

    1. No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

    2. Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**. 

    3. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office. 

    4. Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.

Nesse momento, as funções personalizadas pré-criadas em seu projeto são carregadas e ficam disponíveis no Excel. Experimente a função personalizada `ADD` concluindo as seguintes etapas no Excel:

1. Em uma célula, digite **=CONTOSO**. Observe que o menu de preenchimento automático mostra a lista de todas as funções no namespace `CONTOSO`.

2. Execute a função `CONTOSO.ADD`, com os números `10` e `200` como parâmetros de entrada, especificando o valor a seguir na célula e pressionando Enter:

    ```
    =CONTOSO.ADD(10,200)
    ```

A função personalizada `ADD` calcula a soma de dois números especificados como parâmetros de entrada. Digitar `=CONTOSO.ADD(10,200)` deve produzir o resultado **210** na célula depois que você pressionar Enter.

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Criar uma função personalizada que solicita dados da Web

E se você precisasse de uma função que pudesse solicitar o preço de uma ação a uma API e exibir o resultado na célula de uma planilha? Funções personalizadas são projetadas para que você possa facilmente solicitar os dados da Web de maneira assíncrona.

Conclua as seguintes etapas para criar uma função personalizada denominada `stockPrice`, que aceita um registrador de cotações (por exemplo, **MSFT**) e retorna o preço dessa ação. Essa função personalizada usa a API comercial da IEX, a qual é gratuita e não requer autenticação.

1. No projeto **cotação de ações** que o gerador Yo Office criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.

2. Adicione o código a seguir a **customfunctions.js** e salve o arquivo.

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

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. Antes que o Excel possa disponibilizar essa nova função para os usuários finais, você deve especificar metadados que a descrevam. No projeto **cotação de ações** que o gerador Yo Office criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código. Adicione o objeto a seguir à matriz `functions` do arquivo **config/customfunctions.json** e salve o arquivo.

    Esse JSON descreve a função `stockPrice`.

    ```json
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
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais. Conclua as etapas a seguir para a plataforma que você estiver usando neste tutorial.

    * Se estiver usando o Excel para Windows:

        1. Feche e reabra o Excel.

        2. No Excel, escolha a guia **Inserir**, depois escolha a seta para baixo localizada à direita de **Meus suplementos**.  ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)

        1. Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.
            ![Inserir a faixa de opções no Excel para Windows com o suplemento Funções personalizadas do Excel destacado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)

    * Se estiver usando o Excel Online: 

        1. No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

        2. Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**. 

        3. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office. 

        4. Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.

5. Agora vamos experimentar a nova função. Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione Enter. Você verá que o resultado da célula **B1** é o preço de estoque atual de um compartilhamento de ações da Microsoft.

## <a name="create-a-streaming-asynchronous-custom-function"></a>Criar uma função personalizada assíncrona de fluxo contínuo

A função `stockPrice` que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços de ações estão sempre mudando. Vamos criar uma função personalizada que faça o fluxo de dados de uma API para obter atualizações em tempo real do preço de uma ação.

Conclua as seguintes etapas para criar uma função personalizada denominada `stockPriceStream` que solicita o preço especificado de ações a cada 1.000 milissegundos (desde que a solicitação anterior tenha sido concluída). Enquanto a solicitação inicial estiver em andamento, talvez você veja o valor de espaço reservado **#GETTING_DATA** na célula na qual a função está sendo chamada. Quando um valor é retornado pela função, **#GETTING_DATA** será substituído pelo valor na célula.

1. No projeto **cotação de ações** que o gerador Yo Office criou, adicione código a seguir para **src/customfunctions.js** e salve o arquivo.

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

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. Antes que o Excel possa disponibilizar essa nova função para os usuários finais, você deve especificar metadados que a descrevam. No projeto **cotação de ações** que o gerador Yo Office criou, adicione o seguinte objeto à matriz `functions` do arquivo **config/customfunctions.json** e salve o arquivo.

    Esse JSON descreve a função `stockPriceStream`. Para qualquer função de fluxo contínuo, as propriedades `stream` e `cancelable` devem ser definidas como `true` no objeto `options`, como mostrado neste exemplo de código.

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
                "description": "stock ticker name",
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

3. Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais. Conclua as etapas a seguir para a plataforma que você estiver usando neste tutorial.

    * Se estiver usando o Excel para Windows:

        1. Feche e reabra o Excel.
        
        2. No Excel, escolha a guia **Inserir**, depois escolha a seta para baixo localizada à direita de **Meus suplementos**.  ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)

        3. Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.
            ![Inserir a faixa de opções no Excel para Windows com o suplemento Funções personalizadas do Excel destacado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)

    * Se estiver usando o Excel Online: 

        1. No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

        2. Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**. 

        3. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office. 

        4. Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.

4. Agora vamos experimentar a nova função. Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione Enter. Se o mercado de ações estiver aberto, o resultado na célula **C1** deve ser constantemente atualizado para refletir o preço em tempo real para um compartilhamento de ações da Microsoft.

## <a name="next-steps"></a>Próximas etapas

Neste tutorial, você criou um novo projeto de funções personalizadas, testou uma função pré-criada, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que faz o fluxo de dados em tempo real na Web. Para saber mais sobre as funções personalizadas no Excel, prossiga para o seguinte artigo: 

> [!div class="nextstepaction"]
> [Criar funções personalizadas no Excel](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>Informações legais

Dados fornecidos gratuitamente pela [IEX](https://iextrading.com/developer/). Exibir os [Termos de uso da IEX](https://iextrading.com/api-exhibit-a/). O uso da API da IEX pela Microsoft neste tutorial é apenas para fins educacionais.
