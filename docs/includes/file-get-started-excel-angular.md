# <a name="build-an-excel-add-in-using-angular"></a>Criar um suplemento do Excel usando o Angular

Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o Angular e a API JavaScript do Excel.

## <a name="prerequisites"></a>Pré-requisitos

- Verifique se você já tem os [pré-requisitos de CLI do Angular](https://github.com/angular/angular-cli#prerequisites) e instale todos os pré-requisitos ausentes.

- Instale globalmente a [CLI do Angular](https://github.com/angular/angular-cli). 

    ```bash
    npm install -g @angular/cli
    ```

- Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a>Gerar um novo aplicativo do Angular

Use a CLI do Angular para gerar seu aplicativo do Angular. No terminal, execute o seguinte comando:

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a>Gerar o arquivo de manifesto.

Um arquivo de manifesto do suplemento define seus recursos e configurações.

1. Navegue até a pasta do seu aplicativo.

    ```bash
    cd my-addin
    ```

2. Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento. Execute o comando a seguir e responda aos prompts conforme mostrado abaixo.

    ```bash
    yo office 
    ```

    - **Escolha um tipo de projeto:** `Manifest`
    - **Como deseja nomear seu suplemento?** `My Office Add-in`
    - **Para qual aplicativo cliente do Office você deseja suporte?** `Excel`


    Depois de concluir o assistente, um arquivo de manifesto e um arquivo de recurso estarão disponíveis para você criar o seu projeto.

    ![Gerador do Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > Se for solicitada a substituição de **package.json**, responda **Não** (não substituir).

## <a name="secure-the-app"></a>Proteger o aplicativo

[!include[HTTPS guidance](../includes/https-guidance.md)]

Para este início rápido, é possível usar os certificados fornecidos pelo **Gerador Yeoman para Suplementos do Office**. Você já instalou o gerador globalmente (como parte dos **Pré-requisitos** para este início rápido), então só será preciso copiar os certificados do local de instalação global para a pasta do aplicativo. As etapas a seguir descrevem como concluir esse processo.

1. No terminal, execute o seguinte comando para identificar a pasta onde as bibliotecas globais **npm** estão instaladas:

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > A primeira linha de saída gerada por esse comando especifica a pasta onde as bibliotecas globais **npm** estão instaladas.          
    
2. Usando o Explorador de arquivos, navegue até a pasta `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base`. A partir desse local, copie a pasta `certs` para a área de transferência.

3. Navegue até a pasta raiz do aplicativo Angular que você criou na etapa 1 da seção anterior e cole a pasta `certs` da área de transferência para essa pasta.

## <a name="update-the-app"></a>Atualizar o aplicativo

1. No editor de código, abra **package.json** na raiz do projeto. Modifique o script `start` para especificar que o servidor execute em SSL e porta 3000 e salve o arquivo.

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. Abra **.angular-cli.json** na raiz do projeto. Modifique o objeto **padrões** para especificar o local dos arquivos de certificado e salve o arquivo.

    ```json
    "defaults": {
      "styleExt": "css",
      "component": {},
      "serve": {
        "sslKey": "certs/server.key",
        "sslCert": "certs/server.crt"
      }
    }
    ```

3. Abra **src/index.html**, adicione a marca `<script>` a seguir imediatamente antes da marca `</head>` e salve o arquivo.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. Abra **src/main.ts**, substitua `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` pelo código a seguir e salve o arquivo. 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. Abra **src/polyfills.ts**, adicione a linha de código a seguir acima de todas as outras instruções `import` existentes e salve o arquivo.

    ```typescript
    import 'core-js/client/shim';
    ```

6. No **src/polyfills.ts**, remova a marca de comentário das linhas a seguir e salve o arquivo.

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

7. Abra **src/app/app.component.html**, substitua o conteúdo do arquivo pelo HTML a seguir e salve o arquivo. 

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button (click)="onSetColor()">Set color</button>
        </div>
    </div>
    ```

8. Abra **src/app/app.component.css**, substitua o conteúdo do arquivo pelo código de CSS a seguir e salve o arquivo.

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

9. Abra **src/app/app.component.ts**, substitua o conteúdo do arquivo pelo código a seguir e salve o arquivo. 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onSetColor() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="start-the-dev-server"></a>Iniciar o servidor de desenvolvimento

1. No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.

    ```bash
    npm run start
    ```

2. Em um navegador da web, acesse `https://localhost:3000`. Se o navegador indicar que o certificado do site não é confiável, adicione o certificado como confiável. Veja detalhes em [Adicionar certificados autoassinados como certificados raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

    > [!NOTE]
    > O Chrome (navegador da Web) pode continuar a indicar que o certificado do site não é confiável, mesmo depois de concluir o processo descrito em [Adição de certificados autoassinados como certificado raiz confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Você pode ignorar esse aviso no Chrome e verificar se o certificado é confiável ao navegar até `https://localhost:3000` no Microsoft Edge ou no Internet Explorer. 

3. Depois que o navegador carregar a página do suplemento sem erros de certificado, será possível testar o suplemento. 

## <a name="try-it-out"></a>Experimente

1. Siga as instruções da plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.

    - Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

   
2. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Botão do Suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. Selecione um intervalo de células na planilha.

4. No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Próximas etapas

Você criou com êxito um suplemento do Excel usando o Angular!, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.

> [!div class="nextstepaction"]
> [Tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>Veja também

* [Tutorial de suplemento do Excel](../tutorials/excel-tutorial-create-table.md)
* [Principais conceitos da API JavaScript do Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemplos de código do suplemento do Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Referência da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
