# <a name="build-an-excel-add-in-using-react"></a>Criar um suplemento do Excel usando o React

Neste artigo, voc? passar? pelo processo de criar um suplemento do Excel usando o React e a API JavaScript do Excel.

## <a name="environment"></a>Ambiente

- **?rea de Trabalho do Office**: Verifique se voc? tem a ?ltima vers?o do Office instalada. Comandos de suplemento precisam da compila??o 16.0.6769.0000 ou superior (**16.0.6868.0000** recomendada). Saiba como [Instalar a ?ltima vers?o dos aplicativos do Office](http://aka.ms/latestoffice). 
 
- **Office Online**: N?o h? configura??o adicional. Observe que o suporte para comandos no Office Online para contas de trabalho/escola est? em vers?o pr?via.

## <a name="prerequisites"></a>Pr?-requisitos

- Instale globalmente [Criar aplicativo do React](https://github.com/facebookincubator/create-react-app).

    ```bash
    npm install -g create-react-app
    ```

- Instale a ?ltima vers?o do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a>Gerar um novo aplicativo do React

Use Criar aplicativo do React para gerar seu aplicativo do React. No terminal, execute o seguinte comando:

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a>Gerar o arquivo de manifesto e realizar sideload do suplemento

Cada suplemento requer um arquivo de manifesto para definir os recursos e configura??es.

1. Navegue at? a pasta do seu aplicativo.

    ```bash
    cd my-addin
    ```

2. Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento. Execute o comando a seguir e responda aos prompts, conforme mostrado na seguinte captura de tela:

    ```bash
    yo office
    ```

    - **Gostaria de criar uma nova subpasta para o seu projeto?:** `No`
    - **Como deseja nomear seu suplemento?:** `My Office Add-in`
    - **Para qual aplicativo cliente do Office voc? deseja suporte?:** `Excel`
    - **Gostaria de criar um novo suplemento?:** `No`

    O gerador perguntar? se voc? deseja abrir **resource.html**. N?o ? necess?rio abri-lo para este tutorial, mas fique ? vontade em fazer isso se tiver curiosidade. Escolha Sim ou N?o para concluir o assistente e deixar o gerador fazer seu trabalho.

    ![Gerador do Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > Se for solicitada a substitui??o de **package.json**, responda **N?o** (n?o substituir).

3. Siga as instru??es da plataforma que voc? usar? para executar o suplemento e realizar sideload do suplemento no Excel.

    - Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## <a name="update-the-app"></a>Atualizar o aplicativo

1. Abra **public/index.html**, adicione a marca `<script>` a seguir imediatamente antes da marca `</head>` e salve o arquivo.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. Abra **src/index.js**, substitua `ReactDOM.render(<App />, document.getElementById('root'));` pelo c?digo a seguir e salve o arquivo. 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. Abra **src/App.js**, substitua o conte?do do arquivo pelo c?digo a seguir e salve o arquivo. 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onSetColor = this.onSetColor.bind(this);
      }

      onSetColor() {
        window.Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'green';
          await context.sync();
        });
      }

      render() {
        return (
          <div id="content">
            <div id="content-header">
              <div className="padding">
                  <h1>Welcome</h1>
              </div>
            </div>
            <div id="content-main">
              <div className="padding">
                  <p>Choose the button below to set the color of the selected range to green.</p>
                  <br />
                  <h3>Try it out</h3>
                  <button onClick={this.onSetColor}>Set color</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. Abra **src/App.css**, substitua o conte?do do arquivo pelo c?digo de CSS a seguir e salve o arquivo. 

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

## <a name="try-it-out"></a>Experimente

1. No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.

    Windows:
    ```bash
    set HTTPS=true&&npm start
    ```

    macOS:
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > Uma nova janela de navegador ser? aberta contendo o suplemento. Feche esta janela.

2. No Excel, escolha a guia **P?gina Inicial** e o bot?o **Mostrar Painel de Tarefas** na faixa de op??es para abrir o painel de tarefas do suplemento.

    ![Bot?o do Suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. Selecione um intervalo de c?lulas na planilha.

4. No painel de tarefas, escolha o bot?o **Definir cor** para definir a cor do intervalo selecionado como verde.

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Pr?ximas etapas

Voc? criou com ?xito um suplemento do Excel usando o React, parab?ns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.

> [!div class="nextstepaction"]
> [Tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>Veja tamb?m

* [Tutorial de suplemento do Excel](../tutorials/excel-tutorial-create-table.md)
* [Principais conceitos da API JavaScript do Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemplos de c?digo do suplemento do Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Refer?ncia da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
