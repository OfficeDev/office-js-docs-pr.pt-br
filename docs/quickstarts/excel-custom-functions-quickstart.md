---
ms.date: 03/23/2022
description: Desenvolvendo funções personalizadas no guia de início rápido do Excel.
title: 'Início rápido de funções personalizadas '
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: cac81cb25b9880a3057e2246d39ac226666a4cb4
ms.sourcegitcommit: 64942cdd79d7976a0291c75463d01cb33a8327d8
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/25/2022
ms.locfileid: "64404700"
---
# <a name="get-started-developing-excel-custom-functions"></a>Introdução ao desenvolvimento de funções personalizadas do Excel

Com as funções personalizadas, os desenvolvedores podem adicionar novas funções ao Excel definindo-as em JavaScript ou TypeScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Office conectado a uma assinatura Microsoft 365 (incluindo o Office na web).

  > [!NOTE]
  > Se você ainda não tem o Office, poderá [ingressar no programa para desenvolvedores do Microsoft 365](https://developer.microsoft.com/office/dev-program) para obter uma assinatura do Microsoft 365 gratuita e renovável por 90 dias para usar durante o desenvolvimento.

## <a name="build-your-first-custom-functions-project"></a>Crie seu primeiro projeto com funções personalizadas

Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas. Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Escolha o tipo de projeto:** `Excel Custom Functions Add-in project`
    - **Escolha o tipo de script:** `JavaScript`
    - **Qual será o nome do suplemento?** `starcount`

    :::image type="content" source="../images/starcountPrompt.png" alt-text="Captura de tela da interface de linha de comando do gerador do suplemento Yeoman Office para projetos de funções personalizadas.":::

    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.

1. O gerador Yeoman fornecerá algumas instruções na linha de comando sobre o que fazer com o projeto, mas ignore-as e continue seguindo nossas instruções. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd starcount
    ```

1. Compile o projeto.

    ```command&nbsp;line
    npm run build
    ```

1. Inicie o servidor local da web, que é executado no Node.js. Você pode experimentar o suplemento de função personalizada no Excel. Você pode ser solicitado a abrir o painel de tarefas do suplemento, embora seja opcional. Ainda é possível executar as funções personalizadas sem abrir o painel de tarefas do suplemento.

# <a name="excel-on-windows-or-mac"></a>[Excel para Windows ou Mac](#tab/excel-windows)

Para testar o seu suplemento no Excel para Windows ou Mac, execute o seguinte comando. Quando você executa este comando, o servidor Web local iniciará e o Excel abrirá com o seu suplemento carregado.

```command&nbsp;line
npm run start:desktop
```

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

# <a name="excel-on-the-web"></a>[Excel na Web](#tab/excel-online)

Para testar o suplemento no Excel na Web, execute o seguinte comando. O servidor Web local será iniciado ao executar este comando. Substitua “{url}” pelo URL de um documento do Excel no seu OneDrive ou uma biblioteca do SharePoint para a qual você tenha permissões.

[!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

---

## <a name="try-out-a-prebuilt-custom-function"></a>Experimente uma função personalizada predefinida

O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas definidas no arquivo **./src/functions/functions.js**. O arquivo **./manifest.xml** na pasta raiz do projeto especifica que todas as funções personalizadas pertencem ao namespace `CONTOSO`.

Em sua pasta de trabalho do Excel experimente a função personalizada `ADD` preenchendo as seguintes etapas.

1. Selecione uma célula e um tipo `=CONTOSO`. Observe que o menu de preenchimento automático mostra a lista de todas as funções no namespace `CONTOSO`.

1. Executar a função `CONTOSO.ADD`, usando os números `10` e `200` como parâmetros de entrada, digitando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.

O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada. Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.

[!include[Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

## <a name="next-steps"></a>Próximas etapas

Você criou com êxito uma função personalizada em um suplemento do Excel, parabéns! Em seguida, crie um suplemento mais complexo com o recurso de fluxo de dados. O link a seguir mostra as próximas etapas do tutorial do suplemento do Excel com funções personalizadas.

> [!div class="nextstepaction"]
> [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web)

## <a name="troubleshooting"></a>Solução de problemas

Você poderá encontrar problemas se executar o início rápido várias vezes. Se o Office cache já tiver uma instância de uma função com o mesmo nome, o seu complemento obtém um erro quando ele é sideload. Você pode impedir isso [limpando o cache Office ](../testing/clear-cache.md) antes de executar `npm run start`.

:::image type="content" source="../images/custom-function-already-exists-error.png" alt-text="Uma mensagem de erro Excel intitulada &quot;Funções de instalação de erro&quot;. Ele contém o texto 'Esse complemento não foi instalado porque uma função personalizada com o mesmo nome já existe'.":::

## <a name="see-also"></a>Confira também

- [Visão geral de funções personalizadas](../excel/custom-functions-overview.md)
- [Metadados de funções personalizadas](../excel/custom-functions-json.md)
- [Tempo de execução de funções personalizadas do Excel](../excel/custom-functions-runtime.md)
