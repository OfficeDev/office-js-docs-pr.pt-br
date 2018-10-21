# <a name="tutorial-create-custom-functions-in-excel"></a>Tutorial: Criar funções personalizadas no Excel

## <a name="introduction"></a>Introdução

As funções personalizadas permitem adicionar novas funções ao Excel, definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que qualquer função nativa no Excel, como `SUM()`. Você pode criar funções personalizadas que executam tarefas simples, como cálculos personalizados ou tarefas mais complexas, como a transmissão de dados em tempo real da Web para uma planilha.

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

* A versão mais recente do [Yeoman](http://yeoman.io/) e o [gerador Yo Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando via prompt de comando:

    ```bash
    npm install -g yo generator-office
    ```

* Excel para Windows (build 10827 ou posterior) ou Excel Online

* Faça parte do [programa Office Insider](https://products.office.com/office-insider) (**Insider** level, antigo "Insider Fast")

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

    Após concluir o assistente, o gerador criará os arquivos do projeto e instalará os componentes do nó de suporte. Os arquivos do projeto podem ser encontrados no repositório [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) do GitHub.

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

O projeto de funções personalizadas criado com o gerador Yo Office contém algumas funções personalizadas pré-criados, definidas dentro do arquivo **src/customfunction.js**. O arquivo **manifest.xml** no diretório raiz do projeto especifica que todas as funções personalizadas pertencem ao namespace `CONTOSO`.

Antes de poder usar qualquer uma das funções personalizadas pré-criadas, você deve registrar o suplemento de funções personalizadas no Excel. Para isso, siga as etapas deste tutorial para a plataforma que você vai usar.

* Se for usar o Excel para Windows para testar suas funções personalizadas:

    1. No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)

    2. Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** e selecione o suplemento de **Funções personalizados do Excel** para registrá-lo.  ![Insira a faixa de opções no Excel para Windows com o Suplemento de funções personalizados do Excel realçado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)

* Se for usar o Excel Online para testar suas funções personalizadas: 

    1. No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

    2. Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**. 

    3. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office. 

    4. Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.

Depois disso, as funções personalizadas pré-criadas do seu projeto já estarão carregadas e disponíveis dentro do Excel. Experimente a função personalizada `ADD` seguindo estas no Excel:

1. Dentro de uma célula, digite **= CONTOSO**. Observe que o menu de preenchimento automático mostra a lista de todas as funções no namespace `CONTOSO`.

2. Execute a função `CONTOSO.ADD`, com os números `10` e `200` como parâmetros de entrada, especificando o valor a seguir na célula e pressionando Enter:

    ```
    =CONTOSO.ADD(10,200)
    ```

A função personalizada `ADD` calcula a soma dos dois números especificados por você como parâmetros de entrada. Ao digitar `=CONTOSO.ADD(10,200)` e pressionar Enter, o resultado **210** deve aparecer na célula.

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Criar uma função personalizada que solicita dados da Web

E se você precisar de uma função que solicita o preço de uma ação a uma API e exibe o resultado em uma célula da planilha? Funções personalizadas são projetadas para que você possa facilmente solicitar dados da web de maneira assíncrona.

Complete as etapas a seguir para criar uma função personalizada denominada `stockPrice` que aceita um ticker de ações (como **MSFT**) e retorna o preço da ação. Essa função personalizada usa a API IEX de trading, que é gratuita e não requer autenticação.

1. No projeto **stock-ticker** criado pelo gerador Yo Office, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.

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

3. Para que o Excel possa disponibilizar essa nova função para os usuários finais, você deve primeiro especificar metadados que a descrevem. No projeto **stock-ticker** criado pelo gerador Yo Office, localize o arquivo **config/customfunctions.json** e abra-o no seu editor de código. Adicione o seguinte objeto à matriz `functions` dentro do arquivo **config/customfunctions.json** e salve-o.

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

4. Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais. Conclua as etapas a seguir para a plataforma que estiver usando neste tutorial.

    * Se estiver usando o Excel para Windows:

        1. Feche e reabra o Excel.

        2. No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)

        1. Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** e selecione o suplemento de **Funções personalizados do Excel** para registrá-lo.  ![Insira a faixa de opções no Excel para Windows com o Suplemento de funções personalizados do Excel realçado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)

    * Se estiver usando o Excel Online: 

        1. No Excel Online, escolha a guia **Inserir** e depois escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

        2. Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**. 

        3. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office. 

        4. Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.

5. Agora, vamos experimentar a nova função. Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione Enter. O resultado da célula **B1** deve ser o preço atual de uma ação da Microsoft.

## <a name="create-a-streaming-asynchronous-custom-function"></a>Criar uma função personalizada assíncrona de fluxo contínuo

A função `stockPrice` que você acaba de criar retorna o preço de uma ação em um momento específico, mas os preços de ações estão em constante mudança. Agora, vamos criar uma função personalizada que transmite dados de uma API para obter atualizações do preço de uma ação em tempo real.

Conclua as etapas a seguir para criar uma função personalizada denominada `stockPriceStream` que solicita o preço da ação especificada a cada 1000 milissegundos (desde que a solicitação anterior tenha sido concluída). Enquanto a solicitação inicial estiver em andamento, talvez você veja o valor espaço reservado **#GETTING_DATA** na célula onde a função está sendo chamada. Quando um valor é retornado pela função, **#GETTING_DATA** é substituído por esse valor.

1. No projeto **stock-ticker** criado pelo gerador Yo Office, adicione código a seguir para **src/customfunctions.js** e salve o arquivo.

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

2. Para que o Excel possa disponibilizar essa nova função para os usuários finais, você deve primeiro especificar metadados que a descrevem. No projeto **stock-ticker** criado pelo gerador Yo Office, adicione o objeto a seguir à matriz `functions` no arquivo **config/customfunctions.json** e salve-o.

    Este JSON descreve a função `stockPriceStream`. Para qualquer função de fluxo contínuo, as propriedades `stream` e `cancelable` devem ser definidas como `true` no objeto `options`, como mostrado neste exemplo de código.

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

3. Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais. Conclua as etapas a seguir para a plataforma que estiver usando neste tutorial.

    * Se estiver usando o Excel para Windows:

        1. Feche e reabra o Excel.
        
        2. No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)

        3. Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** e selecione o suplemento de **Funções personalizados do Excel** para registrá-lo.  ![Insira a faixa de opções no Excel para Windows com o Suplemento de funções personalizados do Excel realçado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)

    * Se estiver usando o Excel Online: 

        1. No Excel Online, escolha a guia **Inserir** e depois escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

        2. Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**. 

        3. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office. 

        4. Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.

4. Agora, vamos experimentar a nova função. Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione Enter. Se o mercado de ações estiver aberto, o resultado na célula **C1** será constantemente atualizado para refletir o preço de uma ação da Microsoft em tempo real.

## <a name="next-steps"></a>Próximas etapas

Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função pré-criada, criou uma função personalizada que solicita dados da web e criou uma função personalizada que transmite dados da web em tempo real. Para saber mais sobre as funções personalizadas no Excel, veja o artigo a seguir: 

> [!div class="nextstepaction"]
> [Criar funções personalizadas no Excel](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>Informações jurídicas

Dados fornecidos gratuitamente pelo [IEX](https://iextrading.com/developer/). Verifique os [Termos de uso do IEX](https://iextrading.com/api-exhibit-a/). O uso da API IEX neste tutorial da Microsoft é apenas para fins educacionais.
