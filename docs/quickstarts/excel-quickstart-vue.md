---
title: Criar um suplemento do painel de tarefas do Excel usando o Vue
description: ''
ms.date: 09/18/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: bcd4f84ce6d09db813c643d2cac8fcc5ce5f76c3
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035298"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="3382a-102">Criar um suplemento do painel de tarefas do Excel usando o Vue</span><span class="sxs-lookup"><span data-stu-id="3382a-102">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="3382a-103">Neste artigo, você passará pelo processo de criar um suplemento do painel de tarefas do Excel usando o Vue e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="3382a-103">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3382a-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="3382a-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="3382a-105">Instale a [CLI do Vue](https://cli.vuejs.org/) globalmente.</span><span class="sxs-lookup"><span data-stu-id="3382a-105">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="3382a-106">Gerar um novo aplicativo Vue</span><span class="sxs-lookup"><span data-stu-id="3382a-106">Generate a new Vue app</span></span>

<span data-ttu-id="3382a-p101">Use a CLI do Vue para gerar um novo aplicativo Vue. No terminal, execute o comando a seguir.</span><span class="sxs-lookup"><span data-stu-id="3382a-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command and then answer the prompts as described below.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="3382a-109">Em seguida, selecione a predefinição `default`.</span><span class="sxs-lookup"><span data-stu-id="3382a-109">Then select the `default` preset.</span></span> <span data-ttu-id="3382a-110">Caso seja solicitado a usar o Yarn ou o NPM como um pacote, você poderá escolher qualquer um deles.</span><span class="sxs-lookup"><span data-stu-id="3382a-110">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="3382a-111">Gerar o arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="3382a-111">Generate the manifest file</span></span>

<span data-ttu-id="3382a-112">Cada suplemento requer um arquivo de manifesto para definir os recursos e configurações.</span><span class="sxs-lookup"><span data-stu-id="3382a-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="3382a-113">Navegue até a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="3382a-113">Navigate to your app folder.</span></span>

    ```command&nbsp;line
    cd my-add-in
    ```

2. <span data-ttu-id="3382a-114">Use o gerador Yeoman para gerar o arquivo de manifesto para o seu suplemento executando o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="3382a-114">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > <span data-ttu-id="3382a-115">Ao executar o comando `yo office`, você receberá informações sobre as políticas de coleta de dados de Yeoman e as ferramentas da CLI do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="3382a-115">When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools.</span></span> <span data-ttu-id="3382a-116">Use as informações fornecidas para responder às solicitações como achar melhor.</span><span class="sxs-lookup"><span data-stu-id="3382a-116">Use the information that's provided to respond to the prompts as you see fit.</span></span> <span data-ttu-id="3382a-117">Se você escolher **Sair** em resposta à segunda solicitação, será necessário executar o comando `yo office` novamente quando estiver pronto para criar seu projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="3382a-117">If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.</span></span>

    <span data-ttu-id="3382a-118">Quando solicitado, forneça as seguintes informações para criar seu projeto de suplemento:</span><span class="sxs-lookup"><span data-stu-id="3382a-118">When prompted, provide the following information to create your add-in project:</span></span>

    - <span data-ttu-id="3382a-119">**Escolha o tipo de projeto:** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="3382a-119">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
    - <span data-ttu-id="3382a-120">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="3382a-120">**What do you want to name your add-in?**</span></span> `my-office-add-in`
    - <span data-ttu-id="3382a-121">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="3382a-121">**Which Office client application would you like to support?**</span></span> `Excel`

    ![Gerador do Yeoman](../images/yo-office-manifest-only-vue.png)

<span data-ttu-id="3382a-123">Após concluir o assistente, uma pasta `my-office-add-in` será criada, contendo um arquivo `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="3382a-123">After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="3382a-124">Você usará o manifesto para sideload e testará seu suplemento no final do início rápido.</span><span class="sxs-lookup"><span data-stu-id="3382a-124">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

> [!TIP]
> <span data-ttu-id="3382a-125">Você pode ignorar as orientações da *próximas etapas* fornecidas pelo gerador Yeoman após a criação do projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="3382a-125">You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created.</span></span> <span data-ttu-id="3382a-126">As instruções passo a passo deste artigo fornecem todas as orientações necessárias para concluir este tutorial.</span><span class="sxs-lookup"><span data-stu-id="3382a-126">The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="3382a-127">Proteger o aplicativo</span><span class="sxs-lookup"><span data-stu-id="3382a-127">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="3382a-128">Para habilitar o HTTPS no seu aplicativo, crie um arquivo `vue.config.js` na pasta raiz do projeto Vue com o seguinte conteúdo:</span><span class="sxs-lookup"><span data-stu-id="3382a-128">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a><span data-ttu-id="3382a-129">Atualizar o aplicativo</span><span class="sxs-lookup"><span data-stu-id="3382a-129">Update the app</span></span>

1. <span data-ttu-id="3382a-130">Abra o arquivo `public/index.html` e adicione a seguinte marca `<script>`, imediatamente antes da marca `</head>`:</span><span class="sxs-lookup"><span data-stu-id="3382a-130">Open `public/index.html`, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="3382a-131">Abra `src/main.js` e substitua os conteúdos pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="3382a-131">Open the `src/main.js` file and replace its contents with the following code:</span></span>

   ```js
   import Vue from 'vue';
   import App from './App.vue';

   Vue.config.productionTip = false;

   window.Office.initialize = () => {
     new Vue({
       render: h => h(App)
     }).$mount('#app');
   };
   ```

3. <span data-ttu-id="3382a-132">Abra`src/App.vue` e substitua os conteúdos de arquivo pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="3382a-132">Open the `src/App.vue` file and replace its contents with the following code:</span></span>

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div id="content-main">
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

## <a name="start-the-dev-server"></a><span data-ttu-id="3382a-133">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="3382a-133">Start the dev server</span></span>

1. <span data-ttu-id="3382a-134">No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="3382a-134">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="3382a-135">Em um navegador da web, acesse `https://localhost:3000` (observe o `https`)..</span><span class="sxs-lookup"><span data-stu-id="3382a-135">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="3382a-136">Se o navegador indicar que o certificado do site não é confiável, [configure o computador para confiar no certificado](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span><span class="sxs-lookup"><span data-stu-id="3382a-136">If your browser indicates that the site's certificate is not trusted, you will need to configure your computer to trust the certificate.</span></span>

3. <span data-ttu-id="3382a-137">Quando a página no `https://localhost:3000` estiver em branco e sem erros de certificado, significa que ela está funcionando.</span><span class="sxs-lookup"><span data-stu-id="3382a-137">When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="3382a-138">O Aplicativo Vue é montado após a inicialização do Office, portanto, ele só mostra itens dentro de um ambiente do Excel.</span><span class="sxs-lookup"><span data-stu-id="3382a-138">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="3382a-139">Experimente</span><span class="sxs-lookup"><span data-stu-id="3382a-139">Try it out</span></span>

1. <span data-ttu-id="3382a-140">Siga as instruções da plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="3382a-140">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="3382a-141">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="3382a-141">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="3382a-142">Navegador Web:[Realizar Sideload de Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="3382a-142">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="3382a-143">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="3382a-143">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="3382a-144">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3382a-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="3382a-146">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="3382a-146">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="3382a-147">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="3382a-147">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="3382a-149">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="3382a-149">Next steps</span></span>

<span data-ttu-id="3382a-150">Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel usando o Vue.</span><span class="sxs-lookup"><span data-stu-id="3382a-150">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="3382a-151">Em seguida, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo as etapas deste tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3382a-151">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="3382a-152">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="3382a-152">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="3382a-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="3382a-153">See also</span></span>

* [<span data-ttu-id="3382a-154">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="3382a-154">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="3382a-155">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="3382a-155">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="3382a-156">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="3382a-156">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="3382a-157">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="3382a-157">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
