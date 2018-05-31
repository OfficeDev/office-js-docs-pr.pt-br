# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="946b7-101">Criar um suplemento do Excel usando o Angular</span><span class="sxs-lookup"><span data-stu-id="946b7-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="946b7-102">Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o Angular e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="946b7-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="946b7-103">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="946b7-103">Prerequisites</span></span>

- <span data-ttu-id="946b7-104">Verifique se você já tem os [pré-requisitos de CLI do Angular](https://github.com/angular/angular-cli#prerequisites) e instale todos os pré-requisitos ausentes.</span><span class="sxs-lookup"><span data-stu-id="946b7-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="946b7-105">Instale globalmente a [CLI do Angular](https://github.com/angular/angular-cli).</span><span class="sxs-lookup"><span data-stu-id="946b7-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="946b7-106">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="946b7-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="946b7-107">Gerar um novo aplicativo do Angular</span><span class="sxs-lookup"><span data-stu-id="946b7-107">Generate a new Angular app</span></span>

<span data-ttu-id="946b7-108">Use a CLI do Angular para gerar seu aplicativo do Angular.</span><span class="sxs-lookup"><span data-stu-id="946b7-108">Use the Angular CLI to generate your Angular app.</span></span> <span data-ttu-id="946b7-109">No terminal, execute o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="946b7-109">From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="946b7-110">Gerar o arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="946b7-110">Generate the manifest file</span></span>

<span data-ttu-id="946b7-111">Um arquivo de manifesto do suplemento define seus recursos e configurações.</span><span class="sxs-lookup"><span data-stu-id="946b7-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="946b7-112">Navegue até a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="946b7-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="946b7-113">Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="946b7-113">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="946b7-114">Execute o comando a seguir e responda aos prompts conforme mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="946b7-114">Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office
    ```
    - <span data-ttu-id="946b7-115">**Gostaria de criar uma nova subpasta para o seu projeto?** `No`</span><span class="sxs-lookup"><span data-stu-id="946b7-115">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="946b7-116">**Como deseja nomear seu suplemento?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="946b7-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="946b7-117">**Para qual aplicativo cliente do Office você deseja suporte?** `Excel`</span><span class="sxs-lookup"><span data-stu-id="946b7-117">**Which Office client application would you like to support?:** `Excel`</span></span>
    - <span data-ttu-id="946b7-118">**Gostaria de criar um novo suplemento?** `No`</span><span class="sxs-lookup"><span data-stu-id="946b7-118">**Would you like to create a new add-in?:** `No`</span></span>

    <span data-ttu-id="946b7-p103">O gerador perguntará se você deseja abrir **resource.html**. Não é necessário abri-lo para este tutorial, mas fique à vontade em fazer isso se tiver curiosidade. Escolha Sim ou Não para concluir o assistente e deixar o gerador fazer seu trabalho.</span><span class="sxs-lookup"><span data-stu-id="946b7-p103">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Gerador do Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="946b7-123">Se for solicitada a substituição de **package.json**, responda **Não** (não substituir).</span><span class="sxs-lookup"><span data-stu-id="946b7-123">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="946b7-124">Proteger o aplicativo</span><span class="sxs-lookup"><span data-stu-id="946b7-124">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="946b7-125">Para este início rápido, é possível usar os certificados fornecidos pelo **gerador Yeoman para Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="946b7-125">For this quickstart, you can use the certificates that the **Yeoman generator for Office Add-ins** provides.</span></span> <span data-ttu-id="946b7-126">Você já instalou o gerador globalmente (como parte dos **Pré-requisitos** para este início rápido), portanto só precisa copiar os certificados do local de instalação global para a pasta do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="946b7-126">You've already installed the generator globally (as part of the **Prerequisites** for this quickstart), so you'll just need to copy the certificates from the global install location into your app folder.</span></span> <span data-ttu-id="946b7-127">As etapas a seguir descrevem como concluir esse processo.</span><span class="sxs-lookup"><span data-stu-id="946b7-127">The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="946b7-128">No terminal, execute o seguinte comando para identificar a pasta onde as bibliotecas globais **npm** estão instaladas:</span><span class="sxs-lookup"><span data-stu-id="946b7-128">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="946b7-129">A primeira linha de saída gerada por esse comando especifica a pasta onde as bibliotecas globais **npm** estão instaladas.</span><span class="sxs-lookup"><span data-stu-id="946b7-129">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="946b7-130">Usando o Explorador de arquivos, navegue até a pasta `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base`.</span><span class="sxs-lookup"><span data-stu-id="946b7-130">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder.</span></span> <span data-ttu-id="946b7-131">A partir desse local, copie a pasta `certs` para a área de transferência.</span><span class="sxs-lookup"><span data-stu-id="946b7-131">From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="946b7-132">Navegue até a pasta raiz do aplicativo Angular que você criou na etapa 1 da seção anterior e cole a pasta `certs` da área de transferência para essa pasta.</span><span class="sxs-lookup"><span data-stu-id="946b7-132">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="946b7-133">Atualizar o aplicativo</span><span class="sxs-lookup"><span data-stu-id="946b7-133">Update the app</span></span>

1. <span data-ttu-id="946b7-134">No editor de código, abra **package.json** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="946b7-134">In your code editor, open **package.json** in the root of the project.</span></span> <span data-ttu-id="946b7-135">Modifique o script `start` para especificar que o servidor execute em SSL e porta 3000 e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-135">Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="946b7-136">Abra **.angular-cli.json** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="946b7-136">Open **.angular-cli.json** in the root of the project.</span></span> <span data-ttu-id="946b7-137">Modifique o objeto **padrões** para especificar o local dos arquivos de certificado e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-137">Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

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

3. <span data-ttu-id="946b7-138">Abra **src/index.html**, adicione a marca `<script>` a seguir imediatamente antes da marca `</head>` e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-138">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="946b7-139">Abra **src/main.ts**, substitua `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-139">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="946b7-140">Abra **src/polyfills.ts**, adicione a linha de código a seguir acima de todas as outras instruções `import` existentes e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-140">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="946b7-141">No **src/polyfills.ts**, remova a marca de comentário das linhas a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-141">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

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

7. <span data-ttu-id="946b7-142">Abra **src/app/app.component.html**, substitua o conteúdo do arquivo pelo HTML a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-142">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

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

8. <span data-ttu-id="946b7-143">Abra **src/app/app.component.css**, substitua o conteúdo do arquivo pelo código de CSS a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-143">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

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

9. <span data-ttu-id="946b7-144">Abra **src/app/app.component.ts**, substitua o conteúdo do arquivo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="946b7-144">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

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

## <a name="start-the-dev-server"></a><span data-ttu-id="946b7-145">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="946b7-145">Start the dev server</span></span>

1. <span data-ttu-id="946b7-146">No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="946b7-146">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="946b7-p108">Em um navegador da web, acesse `https://localhost:3000`. Se o navegador indicar que o certificado do site não é confiável, adicione o certificado como confiável. Veja detalhes em [Adicionar certificados autoassinados como certificados raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="946b7-p108">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="946b7-150">O Chrome (navegador da Web) pode continuar a indicar que o certificado do site não é confiável, mesmo depois de concluir o processo descrito em [Adição de certificados autoassinados como certificado raiz confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="946b7-150">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span> <span data-ttu-id="946b7-151">Você pode ignorar esse aviso no Chrome e verificar se o certificado é confiável ao navegar até `https://localhost:3000` no Microsoft Edge ou no Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="946b7-151">You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="946b7-152">Depois que o navegador carregar a página do suplemento sem erros de certificado, será possível testar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="946b7-152">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="946b7-153">Experimente</span><span class="sxs-lookup"><span data-stu-id="946b7-153">Try it out</span></span>

1. <span data-ttu-id="946b7-154">Siga as instruções da plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="946b7-154">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="946b7-155">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="946b7-155">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="946b7-156">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="946b7-156">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="946b7-157">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="946b7-157">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="946b7-158">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="946b7-158">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do Suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="946b7-160">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="946b7-160">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="946b7-161">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="946b7-161">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="946b7-163">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="946b7-163">Next steps</span></span>

<span data-ttu-id="946b7-p110">Você criou com êxito um suplemento do Excel usando o Angular!, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="946b7-p110">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="946b7-166">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="946b7-166">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="946b7-167">Veja também</span><span class="sxs-lookup"><span data-stu-id="946b7-167">See also</span></span>

* [<span data-ttu-id="946b7-168">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="946b7-168">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="946b7-169">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="946b7-169">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="946b7-170">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="946b7-170">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="946b7-171">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="946b7-171">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
