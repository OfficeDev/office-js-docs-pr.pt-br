---
title: Usar o Vue para criar um suplemento do painel de tarefas do Excel
description: Aprenda a criar um suplemento do painel de tarefas simples do Excel usando a API do Office JS e o Vue.
ms.date: 08/04/2021
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 7a463ab61b90914c6fcaebff42599723a3a7e9ee
ms.sourcegitcommit: 8f7d84c33c61c9f724f956740ced01a83f62ddc6
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/01/2022
ms.locfileid: "64605525"
---
# <a name="use-vue-to-build-an-excel-task-pane-add-in"></a>Usar o Vue para criar um suplemento do painel de tarefas do Excel

Neste artigo, você passará pelo processo de criar um suplemento do painel de tarefas do Excel usando o Vue e a API JavaScript do Excel.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Instale a [CLI do Vue](https://cli.vuejs.org/) globalmente. No terminal, execute o seguinte comando.

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a>Gerar um novo aplicativo Vue

Use a CLI do Vue para gerar um novo aplicativo Vue.

```command&nbsp;line
vue create my-add-in
```

Em seguida, selecione a predefinição `Default` para " Vue 3" (se preferir, escolha " Vue 2").

## <a name="generate-the-manifest-file"></a>Gerar o arquivo de manifesto.

Cada suplemento requer um arquivo de manifesto para definir os recursos e configurações.

1. Navegue até a pasta do seu aplicativo.

    ```command&nbsp;line
    cd my-add-in
    ```

1. Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento.

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > Ao executar o comando `yo office`, você receberá informações sobre as políticas de coleta de dados de Yeoman e as ferramentas da CLI do suplemento do Office. Use as informações fornecidas para responder às solicitações como achar melhor. Se você escolher **Sair** em resposta à segunda solicitação, será necessário executar o comando `yo office` novamente quando estiver pronto para criar seu projeto de suplemento.

    Quando solicitado, forneça as informações a seguir para criar seu projeto de suplemento.

    - **Escolha o tipo de projeto:** `Office Add-in project containing the manifest only`
    - **Qual será o nome do suplemento?** `My Office Add-in`
    - **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Excel`

    ![Captura de tela da interface de linha de comando do gerador do Suplemento do Yeoman Office, com o tipo de projeto definido como apenas manifesto.](../images/yo-office-manifest-only-vue.png)

Após a conclusão, o assistente cria uma pasta **Meu suplemento do Office** contendo um arquivo **manifest.xml**. Você usará o manifesto para realizar o sideload e testar o suplemento.

> [!TIP]
> Você pode ignorar as orientações da *próximas etapas* fornecidas pelo gerador Yeoman após a criação do projeto de suplemento. As instruções passo a passo deste artigo fornecem todas as orientações necessárias para concluir este tutorial.

## <a name="secure-the-app"></a>Proteger o aplicativo

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. Habilite HTTPS para seu aplicativo. Na pasta raiz do projeto Vue, crie um arquivo **vue.config.js** com o conteúdo a seguir.

    ```js
    var fs = require("fs");
    var path = require("path");
    var homedir = require('os').homedir()
  
    module.exports = {
      devServer: {
        port: 3000,
        https: {
          key: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.key`)),
          cert: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.crt`)),
          ca: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/ca.crt`)),
         }
       }
    }
    ```

1. Instale os certificados do suplemento.

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador Yeoman contém um código de exemplo para um suplemento básico do painel de tarefas. Se você quiser examinar os principais componentes do seu projeto de suplemento, abra o projeto no seu editor de código e revise os arquivos listados abaixo. Quando estiver pronto para experimentar o suplemento, prossiga para a próxima seção.

- O arquivo **manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento. Para saber mais sobre o arquivo **manifest.xml** arquivo, consulte [manifesto XML de suplementos do Office](../develop/add-in-manifests.md).
- O arquivo **./src/App.vue** contém a marcação HTML para o painel de tarefas, o CSS aplicado ao conteúdo no painel de tarefas e o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o Excel.

## <a name="update-the-app"></a>Atualizar o aplicativo

1. Abra o arquivo **./public/index.html** e adicione o seguinte marca `<script>` antes da marca `</head>`.

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

1. Abra **manifest.xml** e localize as marcas `<bt:Urls>` dentro da marca `<Resources>`. Localize a marca `<bt:Url>` com a ID`Taskpane.Url` e atualize seu atributo `DefaultValue`. O novo `DefaultValue` é `https://localhost:3000/index.html`. Toda a marca atualizada deve corresponder à linha a seguir.

   ```html
   <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/index.html" />
   ```

1. Abra **./src/main.js** e substitua os conteúdos pelo código a seguir.

   ```js
   import { createApp } from 'vue'
   import App from './App.vue'

   window.Office.onReady(() => {
       createApp(App).mount('#app');
   });
   ```

1. Abra **./src/App.vue** e substitua os conteúdos de arquivo pelo código a seguir.

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

## <a name="start-the-dev-server"></a>Inicie o servidor de desenvolvimento

1. Inicie o servidor de desenvolvimento.

   ```command&nbsp;line
   npm run serve
   ```

1. Em um navegador da web, acesse `https://localhost:3000` (observe o `https`). Se a página no `https://localhost:3000` estiver em branco e sem erros de certificado, significa que ela está funcionando. O Aplicativo Vue é montado após a inicialização do Office, portanto, ele só mostra itens dentro de um ambiente do Excel.

## <a name="try-it-out"></a>Experimente

1. Execute o suplemento e o sideload do suplemento no Excel. Siga as instruções para a plataforma que você usará:

   - Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
   - Navegador Web:[Realizar Sideload de Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)
   - iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

1. Abra o painel de tarefas do suplemento no Excel. Na faixa de opções da guia **Página Inicial**, escolha o botão **Mostrar Painel de Tarefas**.

   ![Captura de tela do menu da página inicial do Excel, com o botão Mostrar Painel de Tarefas realçado.](../images/excel-quickstart-addin-2a.png)

1. Selecione um intervalo de células na planilha.

1. Defina a cor do intervalo selecionado como verde. No painel de tarefas do suplemento, escolha o botão **Definir cor**.

   ![Captura de tela do Excel, com o painel de tarefas do suplemento aberto.](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com sucesso um suplemento de painel de tarefas Excel usando a Vue! Em seguida, aprenda mais sobre as capacidades de um suplemento Excel e construa um suplemento mais complexo, seguindo junto com o tutorial do suplemento Excel.

> [!div class="nextstepaction"]
> [Tutorial de suplemento do Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Desenvolver Suplementos do Office](../develop/develop-overview.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](../excel/excel-add-ins-core-concepts.md)
- [Exemplos de código do suplemento do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
