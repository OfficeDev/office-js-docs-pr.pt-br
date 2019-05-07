---
ms.date: 05/02/2019
description: Desenvolvimento de funções personalizadas no guia de início rápido do Excel.
title: Início rápido de funções personalizadas
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8eb2630526ce939273024eebd533bd99fa5e94a1
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619889"
---
# <a name="get-started-developing-excel-custom-functions"></a>Introdução ao desenvolvimento de funções personalizadas do Excel

Com funções personalizadas, os desenvolvedores agora podem adicionar novas funções ao Excel, definindo-as em JavaScript ou typescript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa `SUM()`no Excel, como.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Excel para Windows (versão 1810 64 bits ou posterior) ou o Excel Online

* Ingressar o [programa Office Insider](https://products.office.com/office-insider) (nível**Insider**, anteriormente chamado de "Insider – modo rápido")

## <a name="build-your-first-custom-functions-project"></a>Criar seu primeiro projeto de funções personalizadas

Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas. Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.

1. Em uma pasta de sua preferência, execute o comando a seguir e responda aos prompts da seguinte maneira.

    ```command&nbsp;line
    yo office
    ```

    - **Escolha o tipo de projeto:** `Excel Custom Functions Add-in project (...)`
    - **Escolha o tipo de script:** `JavaScript`
    - **Qual será o nome do suplemento?** `stock-ticker`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/yo-office-excel-cf.png)

    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.

2. O gerador Yeoman fornecerá algumas instruções na linha de comando sobre o que fazer com o projeto, mas ignorará e continuarão seguindo as instruções. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd stock-ticker
    ```

3. Compile o projeto. Isso também instalará os certificados de que seu projeto precisa para que funcionem corretamente. 

    ```command&nbsp;line
    npm run build
    ```

4. Inicie o servidor local da web, que é executado no Node. Você pode experimentar o suplemento função personalizada no Excel para Windows ou no Excel online. Você pode ser solicitado a abrir o painel de tarefas do suplemento, embora isso seja opcional. Você ainda pode executar suas funções personalizadas sem abrir o painel de tarefas do suplemento.

> [!NOTE]
> Os suplementos do Office devem usar HTTPS, não HTTP, mesmo quando você está desenvolvendo. Se você for solicitado a instalar um certificado depois de executar `npm run start:desktop`o, aceite o prompt para instalar o certificado que o gerador Yeoman fornece.

# <a name="excel-for-windowstabexcel-windows"></a>[Excel para Windows](#tab/excel-windows)

Para testar seu suplemento no Excel para Windows, execute o seguinte comando. Quando você executar este comando, o servidor Web local será iniciado e o Excel será aberto com o seu suplemento carregado.

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

Para testar seu suplemento no Excel online, execute o seguinte comando. Quando você executar este comando, o servidor Web local será iniciado.

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> Os suplementos do Office devem usar HTTPS, não HTTP, mesmo quando você está desenvolvendo. Se você for solicitado a instalar um certificado depois de executar `npm run start:web`o, aceite o prompt para instalar o certificado que o gerador Yeoman fornece.

Para usar seu suplemento de funções personalizadas, abra uma nova pasta de trabalho no Excel online. Nesta pasta de trabalho, conclua as seguintes etapas para Sideload seu suplemento.

1. No Excel Online, escolha a guia **inserir** pressione e, em seguida, escolha **suplementos**.

   ![Inserir faixa de opções no Excel online com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)
   
2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

4. Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.

---

## <a name="try-out-a-prebuilt-custom-function"></a>Experimente uma função personalizada predefinida

O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas, definidas no arquivo **./src/Functions/functions.js** . O arquivo **./manifest.xml** no diretório raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.

Na sua pasta de trabalho do Excel, `ADD` Experimente a função personalizada realizando as seguintes etapas:

1. Selecione uma célula e digite `=CONTOSO`. Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.

2. Execute a `CONTOSO.ADD` função, usando números `10` e `200` como parâmetros de entrada, digitando o `=CONTOSO.ADD(10,200)` valor na célula e pressionando ENTER.

O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada. Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito uma função personalizada em um suplemento do Excel! Em seguida, crie um suplemento mais complexo com recurso de dados de streaming. O link a seguir o orienta pelas próximas etapas do tutorial do suplemento do Excel com funções personalizadas.

> [!div class="nextstepaction"]
> [Tutorial de suplemento de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a>Confira também

* [Visão geral de funções personalizadas](../excel/custom-functions-overview.md)
* [Metadados de funções personalizadas](../excel/custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](../excel/custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](../excel/custom-functions-best-practices.md).
