# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="dbf52-101">Criar um suplemento do Excel usando o Angular</span><span class="sxs-lookup"><span data-stu-id="dbf52-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="dbf52-102">Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o Angular e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="dbf52-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="dbf52-103">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="dbf52-103">Prerequisites</span></span>

- <span data-ttu-id="dbf52-104">Verifique se você já tem os [pré-requisitos de CLI do Angular](https://github.com/angular/angular-cli#prerequisites) e instale todos os pré-requisitos que estiver faltando.</span><span class="sxs-lookup"><span data-stu-id="dbf52-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="dbf52-105">Instale globalmente a [CLI do Angular](https://github.com/angular/angular-cli).</span><span class="sxs-lookup"><span data-stu-id="dbf52-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="dbf52-106">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador Yeoman para suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="dbf52-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="dbf52-107">Gerar um novo aplicativo do Angular</span><span class="sxs-lookup"><span data-stu-id="dbf52-107">Generate a new Angular app</span></span>

<span data-ttu-id="dbf52-p101">Use a CLI do Angular para gerar o seu aplicativo Angular. Do terminal, execute o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="dbf52-p101">Use the Angular CLI to generate your Angular app. From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="dbf52-110">Gerar o arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="dbf52-110">Generate the manifest file</span></span>

<span data-ttu-id="dbf52-111">Um arquivo de manifesto do suplemento define seus recursos e configurações.</span><span class="sxs-lookup"><span data-stu-id="dbf52-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="dbf52-112">Navegue até a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="dbf52-p102">Use o gerador Yeoman para gerar o arquivo de manifesto para o seu suplemento. Execute o seguinte comando e responda aos prompts conforme mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-p102">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="dbf52-115">**Escolha um tipo de projeto:** `Office Add-in containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="dbf52-115">**Choose a project type:** `Office Add-in containing the manifest only`</span></span>
    - <span data-ttu-id="dbf52-116">**Como deseja nomear seu suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="dbf52-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="dbf52-117">**Para qual aplicativo cliente do Office você deseja oferecer suporte?** `Excel`</span><span class="sxs-lookup"><span data-stu-id="dbf52-117">**Which Office client application would you like to support?:** `Excel`</span></span>

    <span data-ttu-id="dbf52-118">Depois de concluir o assistente, um arquivo de manifesto e um arquivo de recurso estarão disponíveis para você criar o seu projeto.</span><span class="sxs-lookup"><span data-stu-id="dbf52-118">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>

    ![Gerador Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="dbf52-120">Se for solicitada a substituição de **package.json**, responda **Não** (não substituir).</span><span class="sxs-lookup"><span data-stu-id="dbf52-120">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="dbf52-121">Proteger o aplicativo</span><span class="sxs-lookup"><span data-stu-id="dbf52-121">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="dbf52-p103">Para este início rápido, você pode usar os certificados que o **Yeoman gerador de suplementos do Office** fornece. Você já instalou o gerador de globalmente (como parte dos **pré-requisitos** para este início rápido), portanto você precisará apenas copiar os certificados de global local de instalação para a sua pasta de aplicativos. As etapas a seguir descrevem como concluir este processo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-p103">For this quick start, you can use the certificates that the **Yeoman generator for Office Add-ins** provides. You've already installed the generator globally (as part of the **Prerequisites** for this quick start), so you'll just need to copy the certificates from the global install location into your app folder. The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="dbf52-125">No terminal, execute o seguinte comando para identificar a pasta onde as bibliotecas globais **npm** estão instaladas:</span><span class="sxs-lookup"><span data-stu-id="dbf52-125">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="dbf52-126">A primeira linha de saída gerada por esse comando especifica a pasta onde as bibliotecas globais **npm** estão instaladas.</span><span class="sxs-lookup"><span data-stu-id="dbf52-126">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="dbf52-p104">Usando o Gerenciador de arquivos, navegue até a pasta `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base`. A partir desse local, copie a pasta `certs` para a área de transferência.</span><span class="sxs-lookup"><span data-stu-id="dbf52-p104">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder. From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="dbf52-129">Navegue até a pasta raiz do aplicativo Angular que você criou na etapa 1 da seção anterior e cole a pasta `certs` da área de transferência para essa pasta.</span><span class="sxs-lookup"><span data-stu-id="dbf52-129">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="dbf52-130">Atualizar o aplicativo</span><span class="sxs-lookup"><span data-stu-id="dbf52-130">Update the app</span></span>

1. <span data-ttu-id="dbf52-p105">No seu editor de código, abra **package.json** na raiz do projeto. Modifique o script `start` para especificar que o servidor deve executar usando SSL e a porat 3000 e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-p105">In your code editor, open **package.json** in the root of the project. Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="dbf52-p106">Abra **.angular-cli.json** na raiz do projeto. Modifique o objeto **defaults** para especificar o local dos arquivos de certificado e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-p106">Open **.angular-cli.json** in the root of the project. Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

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

3. <span data-ttu-id="dbf52-135">Abra **src/index.html**, adicione a marca `<script>` imediatamente antes da marca `</head>` e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-135">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="dbf52-136">Abra o **src/main.ts**, substitua `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-136">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="dbf52-137">Abra **src/polyfills.ts**, adicione a linha de código a seguir acima de todas as outras instruções `import` existentes e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-137">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="dbf52-138">Em **src/polyfills.ts**, descomente as linhas a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-138">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

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

7. <span data-ttu-id="dbf52-139">Abra **src/app/app.component.html**, substitua o conteúdo do arquivo pelo HTML a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-139">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

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

8. <span data-ttu-id="dbf52-140">Abra **src/app/app.component.css**, substitua o conteúdo do arquivo pelo código de CSS a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-140">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

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

9. <span data-ttu-id="dbf52-141">Abra **src/app/app.component.ts**, substitua o conteúdo do arquivo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dbf52-141">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

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

## <a name="start-the-dev-server"></a><span data-ttu-id="dbf52-142">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="dbf52-142">Start the dev server</span></span>

1. <span data-ttu-id="dbf52-143">No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="dbf52-143">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="dbf52-p107">Em um navegador da web, acesse `https://localhost:3000`. Se o navegador indicar que o certificado do site não é confiável, você precisará adicionar o certificado como um certificado confiável. Confira detalhes em [Adicionar certificados autoassinados como certificados raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="dbf52-p107">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="dbf52-p108">Chrome (navegador da web) pode continuar para indicar o certificado do site não é confiável, mesmo depois de concluir o processo descrito em [Adicionar certificados autoassinados como certificados raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Você pode ignorar esse aviso no Chrome e pode verificar que o certificado é confiável navegando até `https://localhost:3000` no Internet Explorer ou no Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="dbf52-p108">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="dbf52-149">Depois que o navegador carregar a página do suplemento sem erros de certificado, será possível testar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="dbf52-149">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="dbf52-150">Experimente</span><span class="sxs-lookup"><span data-stu-id="dbf52-150">Try it out</span></span>

1. <span data-ttu-id="dbf52-151">Siga as instruções da plataforma que você usará para executar o suplemento e faça o sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="dbf52-151">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="dbf52-152">Windows: [Fazer o sideload de suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="dbf52-152">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="dbf52-153">Excel Online: [Fazer o sideload dos suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="dbf52-153">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="dbf52-154">iPad e Mac: [Fazer o sideload dos suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="dbf52-154">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="dbf52-155">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="dbf52-155">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="dbf52-157">Selecione qualquer intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="dbf52-157">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="dbf52-158">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="dbf52-158">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="dbf52-160">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="dbf52-160">Next steps</span></span>

<span data-ttu-id="dbf52-p109">Parabéns, você criou um suplemento do Excel usando o Angular com sucesso! Em seguida, aprenda mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo o tutorial do suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="dbf52-p109">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="dbf52-163">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="dbf52-163">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="dbf52-164">Confira também</span><span class="sxs-lookup"><span data-stu-id="dbf52-164">See also</span></span>

* [<span data-ttu-id="dbf52-165">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="dbf52-165">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="dbf52-166">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dbf52-166">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="dbf52-167">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="dbf52-167">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="dbf52-168">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dbf52-168">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
