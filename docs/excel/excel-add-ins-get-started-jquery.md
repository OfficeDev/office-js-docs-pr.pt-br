# <a name="build-an-excel-add-in-using-jquery"></a>Criar um suplemento do Excel usando o jQuery

Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o jQuery e a API JavaScript do Excel.

## <a name="prerequisites"></a>Pré-requisitos

Se ainda não tiver feito anteriormente, instale o [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a>Criar o aplicativo Web

1. Crie uma pasta na sua unidade local e nomeie-a como **my-addin**. Esse é o local em que você criará os arquivos para seu aplicativo.

2. Navegue até a pasta do seu aplicativo.

    ```bash
    cd my-addin
    ```

3. Use o gerador do Yeoman de modo a gerar o arquivo de manifesto para o suplemento. Execute o comando a seguir e responda aos prompts, conforme mostrado na seguinte captura de tela:

    ```bash
    yo office
    ```
    ![Gerador do Yeoman](../../images/yo-office-jquery.png)


4. No editor de código, abra **index.html** na raiz do projeto. Esse arquivo especifica o HTML que será renderizado no painel de tarefas do suplemento. 
 
5. Substitua a marca `header` gerada pela seguinte marcação.
 
    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

6. Substitua a marca `main` gerada pela seguinte marcação e salve o arquivo.

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button class="ms-Button" id="set-color">Set color</button>
        </div>
    </div>
    ```

7. Abra o arquivo **app.js** para especificar o script do suplemento. Substitua a expressão de função imediatamente invocada gerada pelo seguinte código e salve o arquivo.

    ```js
    (function () {
        "use strict";

        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

8. Abra o arquivo **app.css** para especificar os estilos personalizados do suplemento. Substitua o conteúdo (exceto o comentário de direitos autorais) pelo conteúdo a seguir e salve o arquivo.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

## <a name="configure-the-manifest-file-and-sideload-the-add-in"></a>Configurar o arquivo de manifesto e fazer sideload do suplemento

1. Abra o arquivo **my-excel-add-in-manifest.xml** para definir as configurações e os recursos do suplemento. 

2. A marca **ProviderName** tem um valor de espaço reservado. Altere-o para `Microsoft`.

3. O **DefaultValue** da marca **DisplayName** tem um valor de espaço reservado. Altere-o para `A task pane add-in for Excel`. 

4. Salve o arquivo, mas não o feche ainda.

## <a name="configure-to-use-http"></a>Configurar para usar HTTP

Os Suplementos Web do Office devem usar HTTPS, não HTTP, mesmo quando você está desenvolvendo. No entanto, para colocar o suplemento em funcionamento rapidamente, este início rápido usará HTTP. Para ativar esse recurso, siga estas etapas:

1. No arquivo de manifesto **my-office-add-in-manifest.xml**, substitua "https" por "http" em todos os lugares. Em seguida, salve e feche o arquivo.

2. Abra o arquivo **bsconfig.json** na raiz do projeto. Altere o valor da propriedade **https** para `false`. Salve o arquivo.


## <a name="try-it-out"></a>Experimente

1. Siga as instruções para a plataforma que você usará para executar o suplemento e fazer sideload do suplemento no Excel.

    - Windows: [Fazer sideload dos Suplementos do Office para teste no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Fazer sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad e Mac: [Fazer sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. Abra um terminal bash na raiz do projeto e execute o seguinte comando para iniciar o servidor de desenvolvimento.

    ```bash
    npm start
    ```

   > **Observação**: uma nova janela de navegador será aberta contendo o suplemento. Feche esta janela.

3. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Botão do Suplemento do Excel](../../images/excel_quickstart_addin_2a.png)

4. Selecione um intervalo de células na planilha.

5. No painel de tarefas, escolha o botão **Colorir-me** para definir a cor do intervalo selecionado como verde.

    ![Suplemento do Excel](../../images/excel_quickstart_addin_2b.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do Excel usando o jQuery! Em seguida, saiba mais sobre os [principais conceitos](excel-add-ins-core-concepts.md) de criação de suplementos do Excel.

## <a name="additional-resources"></a>Recursos adicionais

* [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
* [Explorar trechos com o Script Lab](https://store.office.com/en-001/app.aspx?assetid=WA104380862&ui=en-US&rs=en-001&ad=US&appredirect=false)
* [Exemplos de código do suplemento do Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Referência da API JavaScript do Excel](../../reference/excel/excel-add-ins-reference-overview.md)
