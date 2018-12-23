# <a name="tutorial-create-custom-functions-in-excel"></a>Tutorial: Criar funções personalizadas no Excel

## <a name="introduction"></a>Introdução

Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`. Você pode criar funções personalizadas que realizam tarefas simples como cálculos personalizados ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.

Neste tutorial, você vai:
> [!div class="checklist"]
> * Criar um projeto de funções personalizadas usando o gerador Yo Office
> * Usar uma função personalizada predefinida para realizar um cálculo simples
> * Criar uma função personalizada que solicita dados da web
> * Criar uma função personalizada que transmite os dados da web em tempo real

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

 Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas. Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.

1. Execute o comando a seguir e responda aos prompts da seguinte forma.

    ```
    yo office
    ```

    * Escolha o tipo de projeto:`Excel Custom Functions Add-in project (...)`

    * Escolha um tipo de script: `JavaScript`

    * Qual será o nome do suplemento? `stock-ticker`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/12-10-fork-cf-pic.jpg)

    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte. Os arquivos do project são provenientes de [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) repositório GitHub.

2. Vá até a pasta do projeto.

    ```
    cd stock-ticker
    ```

3. Confie no certificado autoassinado necessário para executar este projeto. Para obter instruções detalhadas para Windows ou Mac, confira [Adicionando Certificados Autoassinados como Certificado Raiz Confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Crie um projeto.

    ```
    npm run build
    ```

5. Inicie o servidor local da web, que é executado no Node.

    * Se estiver usando o Excel para Windows para testar suas funções personalizadas, execute o seguinte comando para inciar o servidor local da web, inicie o Excel e carregue o suplemento:

        ```
         npm run start
        ```
        Depois de executar esse comando, seu prompt de comando mostrará detalhes sobre o que foi feito, outra janela do npm será aberta mostrando os detalhes da compilação, e o Excel iniciará com o seu suplemento carregado. Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente.

    * Se estiver usando o Excel Online para testar suas funções personalizadas, execute o seguinte comando para inciar o servidor local da web:

        ```
        npm run start-web
        ```

         Depois de executar esse comando, outra janela será aberta mostrando os detalhes da compilação. Para usar suas funções, abra uma nova pasta de trabalho no Office Online.

## <a name="try-out-a-prebuilt-custom-function"></a>Experimente uma função personalizada predefinida

O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas definidas no arquivo **src/customfunction.js**. O arquivo **manifest. XML** na pasta raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.

Em sua pasta de trabalho do Excel experimente a função personalizada`ADD` preenchendo as seguintes etapas no Excel:

1. Em uma célula, digite **= CONTOSO**. Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.

2. Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.

O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada. Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Criar uma função personalizada que solicita dados da web

E se você precisasse de uma função que pode solicitar uma API de preço de uma ação e exibir o resultado na célula de uma planilha? Funções personalizadas são projetadas para que você possa facilmente solicitar os dados da web de forma assíncrona.

Conclua as seguintes etapas para criar uma função personalizada chamada `stockPrice` que aceita um símbolo de cotação da bolsa (por exemplo, **MSFT**) e retorna o preço dessa ação. Esta função personalizada usa IEX Trading API, que é gratuito e não requer autenticação.

1. No projeto**cotações** que o gerador Yeoman criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.

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

3. In **customfunctions.js**, locate the line`CustomFunctionMappings.INCREMENT = increment;`, add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

4. Antes que o Excel possa disponibilizar essa nova função, você deve especificar metadados para descrever a função para o Excel. Abrir o arquivo **config/customfunctions.json**. Adicione o seguinte objeto JSON à matriz 'funções' e salve o arquivo.

    Este JSON descreve a `stockPrice` função.

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
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

5. Você deverá registrar novamente o suplemento no Excel para que a nova função esteja disponível para os usuários finais. Conclua as etapas para a plataforma que você está usando neste tutorial.

    * Se você estiver usando o Excel para Windows:

        1. Feche o Excel e abra novamente o Excel.

        2. No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)

        3. Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.
            ![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)

    * Se você estiver usando o Excel Online:

        1. No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

        2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**. 

        3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman. 

        4. Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.

6. Agora, vamos experimentar a nova função. Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione enter. Você verá que o resultado na célula **B1** é o preço atual das ações para uma ação da Microsoft.

## <a name="create-a-streaming-asynchronous-custom-function"></a>Criar uma função personalizada assíncrona de streaming

A `stockPrice` função que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços das ações estão sempre mudando. Vamos criar uma função personalizada de fluxos de dados de uma API recebendo atualizações em tempo real sobre o preço de uma atuação.

Conclua as seguintes etapas para criar uma função personalizada chamada `stockPriceStream` que solicita o preço da ação a cada 1000 milissegundos (desde que a solicitação anterior esteja concluída). Enquanto a solicitação inicial está em andamento, você poderá ver o valor de espaço reservado **# OBTENDO_DADOS** na célula em que a função está sendo exibida. Quando um valor é retornado pela função, **# OBTENDO_DADOS**será substituído por esse valor na célula.

1. No projeto**cotações** que o gerador Yeoman criou, adicione o código a seguir **src/customfunctions.js** e salve o arquivo.

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

2. Antes que o Excel possa fazer esta nova função nova disponível para usuários, especifique os metadados que descreve essa função. No projeto**cotações** que o gerador Yeoman criou, adicione o objeto a seguir na `functions`matriz em **config/customfunctions.json** e salve o arquivo.

    Este JSON descreve a `stockPriceStream` função. Para qualquer função streaming a propriedade `stream` e a propriedade `cancelable` devem ser definidas como `true` dentro do `options` objeto, como mostra este exemplo código.

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

3. Você deverá registrar novamente o suplemento no Excel para que a nova função esteja disponível para os usuários finais. Conclua as etapas para a plataforma que você está usando neste tutorial.

    * Se você estiver usando o Excel para Windows:

        1. Feche o Excel e abra novamente o Excel.
        
        2. No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)

        3. Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.
            ![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)

    * Se você estiver usando o Excel Online:

        1. No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

        2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

        3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

        4. Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.

4. Agora, vamos experimentar a nova função. Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione enter. Desde que o mercado de ações esteja aberto, você verá que o resultado na célula **C1** é constantemente atualizado para refletir o preço em tempo uma ação das ações da Microsoft.

## <a name="next-steps"></a>Próximas etapas

Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função predefinida, criou uma função personalizada que solicita dados da web e criou uma função personalizada que transmite dados em tempo real da Web. Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:

> [!div class="nextstepaction"]
> [Criar funções personalizadas no Excel](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>Informações legais

Dados gratuito fornecidos pela [IEX](https://iextrading.com/developer/). Modo de exibição [termos de uso IEX](https://iextrading.com/api-exhibit-a/). O uso da Microsoft dA API IEX neste tutorial é apenas para fins educacionais.
