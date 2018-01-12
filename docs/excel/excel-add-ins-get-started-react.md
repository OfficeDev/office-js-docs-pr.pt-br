# <a name="build-an-excel-add-in-using-react"></a>Criar um suplemento do Excel usando o React

Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o React e a API JavaScript do Excel.

## <a name="prerequisites"></a>Pré-requisitos

Se ainda não tiver feito isso, será necessário instalar as seguintes ferramentas:

1. Instale [Criar Aplicativo React](https://github.com/facebookincubator/create-react-app) globalmente.

    ```bash
    npm install -g create-react-app
    ```

2. Instale o [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a>Gerar um novo aplicativo do React

Use Criar aplicativo do React para gerar seu aplicativo do React. No terminal, execute o seguinte comando:

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a>Gerar o arquivo de manifesto e fazer sideload do suplemento

Cada suplemento requer um arquivo de manifesto para definir seus recursos e configurações.

1. Navegue até a pasta do seu aplicativo.

    ```bash
    cd my-addin
    ```

2. Use o gerador do Yeoman de modo a gerar o arquivo de manifesto para o suplemento. Execute o comando a seguir e responda aos prompts, conforme mostrado na seguinte captura de tela:

    ```bash
    yo office
    ```
    ![Gerador do Yeoman](../../images/yo-office.png)
    >**Observação**: Se for solicitada a substituição de **package.json**, responda **Não** (não substituir).

3. Abra o arquivo de manifesto (isto é, o arquivo no diretório raiz do aplicativo com um nome que termina em "manifest.xml"). Substitua todas as ocorrências de `https://localhost:3000` por `http://localhost:3000` e salve o arquivo.

4. Siga as instruções para a plataforma que você usará para executar o suplemento e fazer sideload do suplemento no Excel.

    - Windows: [Fazer sideload dos Suplementos do Office para teste no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Fazer sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad e Mac: [Fazer sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## <a name="update-the-app"></a>Atualizar o aplicativo

1. Abra **public/index.html**, adicione a marca `<script>` a seguir imediatamente antes da marca `</head>` e salve o arquivo.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. Abra **src/index.js**, substitua `ReactDOM.render(<App />, document.getElementById('root'));` pelo código a seguir e salve o arquivo. 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. Abra **src/App.js**, substitua o conteúdo do arquivo pelo código a seguir e salve o arquivo. 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onColorMe = this.onColorMe.bind(this);
      }

      onColorMe() {
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
                  <button onClick={this.onColorMe}>Color Me</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. Abra **src/App.css**, substitua o conteúdo do arquivo pelo código de CSS a seguir e salve o arquivo. 

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

    ```bash
    npm start
    ```

2. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Botão do Suplemento do Excel](../../images/excel_quickstart_addin_2a.png)

3. No painel de tarefas, escolha o botão **Colorir-me** para definir a cor do intervalo selecionado como verde.

    ![Suplemento do Excel](../../images/excel_quickstart_addin_2b.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do Excel usando o React! Em seguida, saiba mais sobre os [principais conceitos](excel-add-ins-core-concepts.md) de criação de suplementos do Excel.

## <a name="additional-resources"></a>Recursos adicionais

* [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
* [Exemplos de código do suplemento do Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Referência da API JavaScript do Excel](../../reference/excel/excel-add-ins-reference-overview.md)
