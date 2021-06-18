---
title: Criar um suplemento do painel de tarefas do Excel usando o Vue
description: Aprenda a criar um suplemento do painel de tarefas simples do Excel usando a API do Office JS e o Vue.
ms.date: 06/16/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: cd709910c9e69478c953c03b5e17d5512e875d91
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007815"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="f2f4e-103">Criar um suplemento do painel de tarefas do Excel usando o Vue</span><span class="sxs-lookup"><span data-stu-id="f2f4e-103">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="f2f4e-104">Neste artigo, você passará pelo processo de criar um suplemento do painel de tarefas do Excel usando o Vue e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-104">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f2f4e-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="f2f4e-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="f2f4e-106">Instale a [CLI do Vue](https://cli.vuejs.org/) globalmente.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-106">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="f2f4e-107">Gerar um novo aplicativo Vue</span><span class="sxs-lookup"><span data-stu-id="f2f4e-107">Generate a new Vue app</span></span>

<span data-ttu-id="f2f4e-p101">Use a CLI do Vue para gerar um novo aplicativo Vue. No terminal, execute o comando a seguir.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="f2f4e-110">Em seguida, selecionar o `Default` predefinido para "Vue 3" (você pode escolher usar "Vue 2", se preferir).</span><span class="sxs-lookup"><span data-stu-id="f2f4e-110">Then select the `Default` preset for "Vue 3" (you may choose to use "Vue 2" if you'd prefer).</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="f2f4e-111">Gerar o arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-111">Generate the manifest file</span></span>

<span data-ttu-id="f2f4e-112">Cada suplemento requer um arquivo de manifesto para definir os recursos e configurações.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="f2f4e-113">Navegue até a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-113">Navigate to your app folder.</span></span>

    ```command&nbsp;line
    cd my-add-in
    ```

2. <span data-ttu-id="f2f4e-114">Use o gerador Yeoman para gerar o arquivo de manifesto para o seu suplemento executando o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="f2f4e-114">Use the Yeoman generator to generate the manifest file for your add-in by running the following command:</span></span>

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > <span data-ttu-id="f2f4e-115">Ao executar o comando `yo office`, você receberá informações sobre as políticas de coleta de dados de Yeoman e as ferramentas da CLI do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-115">When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools.</span></span> <span data-ttu-id="f2f4e-116">Use as informações fornecidas para responder às solicitações como achar melhor.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-116">Use the information that's provided to respond to the prompts as you see fit.</span></span> <span data-ttu-id="f2f4e-117">Se você escolher **Sair** em resposta à segunda solicitação, será necessário executar o comando `yo office` novamente quando estiver pronto para criar seu projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-117">If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.</span></span>

    <span data-ttu-id="f2f4e-118">Quando solicitado, forneça as seguintes informações para criar seu projeto de suplemento:</span><span class="sxs-lookup"><span data-stu-id="f2f4e-118">When prompted, provide the following information to create your add-in project:</span></span>

    - <span data-ttu-id="f2f4e-119">**Escolha o tipo de projeto:** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="f2f4e-119">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
    - <span data-ttu-id="f2f4e-120">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="f2f4e-120">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="f2f4e-121">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="f2f4e-121">**Which Office client application would you like to support?**</span></span> `Excel`

    ![Captura de tela da interface de linha de comando do gerador do Suplemento do Yeoman Office, com o tipo de projeto definido como apenas manifesto](../images/yo-office-manifest-only-vue.png)

<span data-ttu-id="f2f4e-123">Após concluir o assistente, uma pasta `My Office Add-in` será criada, contendo um arquivo `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-123">After you complete the wizard, it creates a `My Office Add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="f2f4e-124">Você usará o manifesto para sideload e testará seu suplemento no final do início rápido.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-124">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

> [!TIP]
> <span data-ttu-id="f2f4e-125">Você pode ignorar as orientações da *próximas etapas* fornecidas pelo gerador Yeoman após a criação do projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-125">You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created.</span></span> <span data-ttu-id="f2f4e-126">As instruções passo a passo deste artigo fornecem todas as orientações necessárias para concluir este tutorial.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-126">The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="f2f4e-127">Proteger o aplicativo</span><span class="sxs-lookup"><span data-stu-id="f2f4e-127">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. <span data-ttu-id="f2f4e-128">Para habilitar o HTTPS no seu aplicativo, crie um arquivo `vue.config.js` na pasta raiz do projeto Vue com o seguinte conteúdo:</span><span class="sxs-lookup"><span data-stu-id="f2f4e-128">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

    ```js
    var fs = require("fs");
    var path = require("path");
    var homedir = require('os').homedir()
  
    module.exports = {
      devServer: {
        port: 3000,
        https: true,
        key: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.key`)),
        cert: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.crt`)),
        ca: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/ca.crt`))
      }
    }
    ```

2. <span data-ttu-id="f2f4e-129">No terminal, execute o seguinte comando para instalar os certificados do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-129">From the terminal, run the following command to install the add-in's certificates.</span></span>

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="update-the-app"></a><span data-ttu-id="f2f4e-130">Atualizar o aplicativo</span><span class="sxs-lookup"><span data-stu-id="f2f4e-130">Update the app</span></span>

1. <span data-ttu-id="f2f4e-131">Abra o arquivo `public/index.html` e adicione a seguinte marca `<script>`, imediatamente antes da marca `</head>`:</span><span class="sxs-lookup"><span data-stu-id="f2f4e-131">Open the `public/index.html` file and add the following `<script>` tag immediately before the `</head>` tag:</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="f2f4e-132">Abra `src/main.js` e substitua os conteúdos pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="f2f4e-132">Open `src/main.js` and replace the contents with the following code:</span></span>

   ```js
   import { createApp } from 'vue'
   import App from './App.vue'

   window.Office.onReady(() => {
       createApp(App).mount('#app');
   });
   ```

3. <span data-ttu-id="f2f4e-133">Abra`src/App.vue` e substitua os conteúdos de arquivo pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="f2f4e-133">Open `src/App.vue` and replace the file contents with the following code:</span></span>

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div class="content-main">
           <div class="padding">
             <p>
               Choose the button below to set the color of the selected range to
               green.
             </p>
             <br />
             <h3>Try it out</h3>
             <button @click="onSetColor">Set color</button>
           </div>
         </div>
       </div>
     </div>
   </template>

   <script>
     export default {
       name: 'App',
       methods: {
         onSetColor() {
           window.Excel.run(async context => {
             const range = context.workbook.getSelectedRange();
             range.format.fill.color = 'green';
             await context.sync();
           });
         }
       }
     };
   </script>

   <style>
     .content-header {
       background: #2a8dd4;
       color: #fff;
       position: absolute;
       top: 0;
       left: 0;
       width: 100%;
       height: 80px;
       overflow: hidden;
     }

     .content-main {
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
   </style>
   ```

## <a name="start-the-dev-server"></a><span data-ttu-id="f2f4e-134">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="f2f4e-134">Start the dev server</span></span>

1. <span data-ttu-id="f2f4e-135">No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-135">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="f2f4e-136">Em um navegador da web, acesse `https://localhost:3000` (observe o `https`)..</span><span class="sxs-lookup"><span data-stu-id="f2f4e-136">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="f2f4e-137">Se a página no `https://localhost:3000` estiver em branco e sem erros de certificado, significa que ela está funcionando.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-137">If the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="f2f4e-138">O Aplicativo Vue é montado após a inicialização do Office, portanto, ele só mostra itens dentro de um ambiente do Excel.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-138">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="f2f4e-139">Experimente</span><span class="sxs-lookup"><span data-stu-id="f2f4e-139">Try it out</span></span>

1. <span data-ttu-id="f2f4e-140">Siga as instruções da plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-140">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="f2f4e-141">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="f2f4e-141">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="f2f4e-142">Navegador Web:[Realizar Sideload de Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="f2f4e-142">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="f2f4e-143">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="f2f4e-143">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="f2f4e-144">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Captura de tela do menu da página inicial do Excel, com o botão Mostrar Painel de Tarefas realçado](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="f2f4e-146">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-146">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="f2f4e-147">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-147">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Captura de tela do Excel, com o painel de tarefas do suplemento aberto](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="f2f4e-149">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="f2f4e-149">Next steps</span></span>

<span data-ttu-id="f2f4e-p106">Parabéns, você criou com sucesso um suplemento de painel de tarefas Excel usando a Vue! Em seguida, aprenda mais sobre as capacidades de um suplemento Excel e construa um suplemento mais complexo, seguindo junto com o tutorial do suplemento Excel.</span><span class="sxs-lookup"><span data-stu-id="f2f4e-p106">Congratulations, you've successfully created an Excel task pane add-in using Vue! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f2f4e-152">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="f2f4e-152">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="f2f4e-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="f2f4e-153">See also</span></span>

* [<span data-ttu-id="f2f4e-154">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f2f4e-154">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="f2f4e-155">Desenvolver Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f2f4e-155">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="f2f4e-156">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f2f4e-156">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="f2f4e-157">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="f2f4e-157">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="f2f4e-158">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f2f4e-158">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
