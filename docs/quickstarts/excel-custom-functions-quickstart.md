---
ms.date: 08/04/2021
description: Desenvolvendo funções personalizadas no guia de início rápido do Excel.
title: 'Início rápido de funções personalizadas '
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: f503948b8f468234221ee8713ea4654a7cc76c07
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152017"
---
# <a name="get-started-developing-excel-custom-functions"></a>Introdução ao desenvolvimento de funções personalizadas do Excel

Com as funções personalizadas, os desenvolvedores agora podem adicionar novas funções ao Excel definindo-as em JavaScript ou Typescript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Excel no Windows (versão 1904 ou posterior, conectado a uma assinatura do Microsoft 365) ou Excel na web
- O Office no Mac (conectado a uma assinatura do Microsoft 365) é compatível com as funções personalizadas do Excel) e uma atualização desse tutorial está a caminho.

>[!NOTE]
>As funções personalizadas do Excel não são compatíveis com o Office 2019 (compra única).

## <a name="build-your-first-custom-functions-project"></a>Crie seu primeiro projeto com funções personalizadas

Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas. Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Escolha o tipo de projeto:** `Excel Custom Functions Add-in project`
    - **Escolha o tipo de script:** `JavaScript`
    - **Qual será o nome do suplemento?** `starcount`

    ![Captura de tela da interface de linha de comando do gerador do suplemento Yeoman Office para projetos de funções personalizadas.](../images/starcountPrompt.png)

    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.

1. O gerador Yeoman fornecerá algumas instruções na linha de comando sobre o que fazer com o projeto, mas ignore-as e continue seguindo nossas instruções. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd starcount
    ```

1. Compile o projeto.

    ```command&nbsp;line
    npm run build
    ```

1. Inicie o servidor local da web, que é executado no Node.js. Você pode experimentar o suplemento função personalizada no Excel na Web ou no Windows. Você pode ser solicitado a abrir o painel de tarefas do suplemento, embora seja opcional. Ainda é possível executar as funções personalizadas sem abrir o painel de tarefas do suplemento.

# <a name="excel-on-windows"></a>[Excel no Windows](#tab/excel-windows)

Para testar o suplemento no Excel para Windows ou Mac, execute o seguinte comando. Quando você executa este comando, o servidor Web local iniciará e o Excel abrirá com o seu suplemento carregado.

```command&nbsp;line
npm run start:desktop
```

> [!NOTE]
> Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se você for solicitado a instalar um certificado após executar `npm run start`, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.
    
# <a name="excel-on-the-web"></a>[Excel na Web](#tab/excel-online)

Para testar o suplemento no Excel na Web, execute o seguinte comando. O servidor Web local será iniciado ao executar este comando.

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se você for solicitado a instalar um certificado após executar `npm run start`, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

Para usar o suplemento de funções personalizadas, abra uma nova pasta de trabalho no Excel em um navegador. Nesta pasta de trabalho, conclua as seguintes etapas para realizar o sideload do suplemento.

1. No Excel, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.

   ![Captura de tela da faixa de opções Inserir no Excel na web, com o botão Meus suplementos destacado.](../images/excel-cf-online-register-add-in-1.png)

1. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

1. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

1. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

---

## <a name="try-out-a-prebuilt-custom-function"></a>Experimente uma função personalizada predefinida

O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas definidas no arquivo **./src/functions/functions.js**. O arquivo **./manifest.xml** na pasta raiz do projeto especifica que todas as funções personalizadas pertencem ao namespace `CONTOSO`.

Em sua pasta de trabalho do Excel experimente a função personalizada `ADD` preenchendo as seguintes etapas.

1. Selecione uma célula e digite `=CONTOSO` Observe que o menu de preenchimento automático mostra a lista de todas as funções no namespace `CONTOSO`.

1. Executar a função `CONTOSO.ADD`, usando os números `10` e `200` como parâmetros de entrada, digitando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.

O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada. Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.

## <a name="next-steps"></a>Próximas etapas

Você criou com êxito uma função personalizada em um suplemento do Excel, parabéns! Em seguida, crie um suplemento mais complexo com o recurso de fluxo de dados. O link a seguir mostra as próximas etapas do tutorial do suplemento do Excel com funções personalizadas.

> [!div class="nextstepaction"]
> [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web)

## <a name="see-also"></a>Confira também

- [Visão geral de funções personalizadas](../excel/custom-functions-overview.md)
- [Metadados de funções personalizadas](../excel/custom-functions-json.md)
- [Tempo de execução de funções personalizadas do Excel](../excel/custom-functions-runtime.md)
