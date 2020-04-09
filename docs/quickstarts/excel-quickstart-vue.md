---
title: Criar um suplemento do painel de tarefas do Excel usando o Vue
description: Aprenda a criar um suplemento do painel de tarefas simples do Excel usando a API do Office JS e o Vue.
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aff58bf3021be2efed0aef14a505dab8433d92a3
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185558"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="3e28e-103">Criar um suplemento do painel de tarefas do Excel usando o Vue</span><span class="sxs-lookup"><span data-stu-id="3e28e-103">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="3e28e-104">Neste artigo, você passará pelo processo de criar um suplemento do painel de tarefas do Excel usando o Vue e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="3e28e-104">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3e28e-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="3e28e-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="3e28e-106">Instale a [CLI do Vue](https://cli.vuejs.org/) globalmente.</span><span class="sxs-lookup"><span data-stu-id="3e28e-106">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="3e28e-107">Gerar um novo aplicativo Vue</span><span class="sxs-lookup"><span data-stu-id="3e28e-107">Generate a new Vue app</span></span>

<span data-ttu-id="3e28e-p101">Use a CLI do Vue para gerar um novo aplicativo Vue. No terminal, execute o comando a seguir.</span><span class="sxs-lookup"><span data-stu-id="3e28e-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="3e28e-110">Em seguida, selecione a predefinição `default`.</span><span class="sxs-lookup"><span data-stu-id="3e28e-110">Then select the `default` preset.</span></span> <span data-ttu-id="3e28e-111">Caso seja solicitado a usar o Yarn ou o NPM como um pacote, você poderá escolher qualquer um deles.</span><span class="sxs-lookup"><span data-stu-id="3e28e-111">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="3e28e-112">Gerar o arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="3e28e-112">Generate the manifest file</span></span>

<span data-ttu-id="3e28e-113">Cada suplemento requer um arquivo de manifesto para definir os recursos e configurações.</span><span class="sxs-lookup"><span data-stu-id="3e28e-113">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="3e28e-114">Navegue até a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="3e28e-114">Navigate to your app folder.</span></span>

    ```command&nbsp;line
    cd my-add-in
    ```

2. <span data-ttu-id="3e28e-115">Use o gerador Yeoman para gerar o arquivo de manifesto para o seu suplemento executando o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="3e28e-115">Use the Yeoman generator to generate the manifest file for your add-in by running the following command:</span></span>

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > <span data-ttu-id="3e28e-116">Ao executar o comando `yo office`, você receberá informações sobre as políticas de coleta de dados de Yeoman e as ferramentas da CLI do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="3e28e-116">When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools.</span></span> <span data-ttu-id="3e28e-117">Use as informações fornecidas para responder às solicitações como achar melhor.</span><span class="sxs-lookup"><span data-stu-id="3e28e-117">Use the information that's provided to respond to the prompts as you see fit.</span></span> <span data-ttu-id="3e28e-118">Se você escolher **Sair** em resposta à segunda solicitação, será necessário executar o comando `yo office` novamente quando estiver pronto para criar seu projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="3e28e-118">If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.</span></span>

    <span data-ttu-id="3e28e-119">Quando solicitado, forneça as seguintes informações para criar seu projeto de suplemento:</span><span class="sxs-lookup"><span data-stu-id="3e28e-119">When prompted, provide the following information to create your add-in project:</span></span>

    - <span data-ttu-id="3e28e-120">**Escolha o tipo de projeto:** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="3e28e-120">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
    - <span data-ttu-id="3e28e-121">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="3e28e-121">**What do you want to name your add-in?**</span></span> `my-office-add-in`
    - <span data-ttu-id="3e28e-122">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="3e28e-122">**Which Office client application would you like to support?**</span></span> `Excel`

    ![Gerador do Yeoman](../images/yo-office-manifest-only-vue.png)

<span data-ttu-id="3e28e-124">Após concluir o assistente, uma pasta `my-office-add-in` será criada, contendo um arquivo `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="3e28e-124">After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="3e28e-125">Você usará o manifesto para sideload e testará seu suplemento no final do início rápido.</span><span class="sxs-lookup"><span data-stu-id="3e28e-125">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

> [!TIP]
> <span data-ttu-id="3e28e-126">Você pode ignorar as orientações da *próximas etapas* fornecidas pelo gerador Yeoman após a criação do projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="3e28e-126">You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created.</span></span> <span data-ttu-id="3e28e-127">As instruções passo a passo deste artigo fornecem todas as orientações necessárias para concluir este tutorial.</span><span class="sxs-lookup"><span data-stu-id="3e28e-127">The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="3e28e-128">Proteger o aplicativo</span><span class="sxs-lookup"><span data-stu-id="3e28e-128">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="3e28e-129">Para habilitar o HTTPS no seu aplicativo, crie um arquivo `vue.config.js` na pasta raiz do projeto Vue com o seguinte conteúdo:</span><span class="sxs-lookup"><span data-stu-id="3e28e-129">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a><span data-ttu-id="3e28e-130">Atualizar o aplicativo</span><span class="sxs-lookup"><span data-stu-id="3e28e-130">Update the app</span></span>

1. <span data-ttu-id="3e28e-131">Abra o arquivo `public/index.html` e adicione a seguinte marca `<script>`, imediatamente antes da marca `</head>`:</span><span class="sxs-lookup"><span data-stu-id="3e28e-131">Open the `public/index.html` file and add the following `<script>` tag immediately before the `</head>` tag:</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="3e28e-132">Abra `src/main.js` e substitua os conteúdos pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="3e28e-132">Open `src/main.js` and replace the contents with the following code:</span></span>

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

3. <span data-ttu-id="3e28e-133">Abra`src/App.vue` e substitua os conteúdos de arquivo pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="3e28e-133">Open `src/App.vue` and replace the file contents with the following code:</span></span>

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

## <a name="start-the-dev-server"></a><span data-ttu-id="3e28e-134">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="3e28e-134">Start the dev server</span></span>

1. <span data-ttu-id="3e28e-135">No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="3e28e-135">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="3e28e-136">Em um navegador da web, acesse `https://localhost:3000` (observe o `https`)..</span><span class="sxs-lookup"><span data-stu-id="3e28e-136">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="3e28e-137">Se o navegador indicar que o certificado do site não é confiável, [configure o computador para confiar no certificado](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span><span class="sxs-lookup"><span data-stu-id="3e28e-137">If your browser indicates that the site's certificate is not trusted, you will need to [configure your computer to trust the certificate](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span></span>

3. <span data-ttu-id="3e28e-138">Quando a página no `https://localhost:3000` estiver em branco e sem erros de certificado, significa que ela está funcionando.</span><span class="sxs-lookup"><span data-stu-id="3e28e-138">When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="3e28e-139">O Aplicativo Vue é montado após a inicialização do Office, portanto, ele só mostra itens dentro de um ambiente do Excel.</span><span class="sxs-lookup"><span data-stu-id="3e28e-139">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="3e28e-140">Experimente</span><span class="sxs-lookup"><span data-stu-id="3e28e-140">Try it out</span></span>

1. <span data-ttu-id="3e28e-141">Siga as instruções da plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="3e28e-141">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="3e28e-142">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="3e28e-142">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="3e28e-143">Navegador Web:[Realizar Sideload de Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="3e28e-143">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="3e28e-144">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="3e28e-144">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="3e28e-145">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3e28e-145">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="3e28e-147">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="3e28e-147">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="3e28e-148">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="3e28e-148">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="3e28e-150">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="3e28e-150">Next steps</span></span>

<span data-ttu-id="3e28e-151">Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel usando o Vue.</span><span class="sxs-lookup"><span data-stu-id="3e28e-151">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="3e28e-152">Em seguida, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo as etapas deste tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3e28e-152">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="3e28e-153">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="3e28e-153">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="3e28e-154">Confira também</span><span class="sxs-lookup"><span data-stu-id="3e28e-154">See also</span></span>

* [<span data-ttu-id="3e28e-155">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3e28e-155">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="3e28e-156">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="3e28e-156">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="3e28e-157">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="3e28e-157">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="3e28e-158">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="3e28e-158">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="3e28e-159">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="3e28e-159">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="3e28e-160">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="3e28e-160">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
