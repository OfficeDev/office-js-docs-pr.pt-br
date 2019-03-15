---
ms.date: 03/06/2019
description: Desenvolvimento de funções personalizadas no guia de início rápido do Excel.
title: Início rápido de funções personalizadas (visualização)
localization_priority: Normal
ms.openlocfilehash: 9dd3e5a99f08ce0b931e705fac3312ab10c19e18
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632699"
---
# <a name="get-started-developing-excel-custom-functions"></a>Introdução ao desenvolvimento de funções personalizadas do Excel

Com funções personalizadas, os desenvolvedores agora podem adicionar novas funções ao Excel, definindo-as em JavaScript ou typescript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa `SUM()`no Excel, como.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Você precisará das seguintes ferramentas e recursos relacionados para começar a criar funções personalizadas.

- [Node](https://nodejs.org/en/) (versão 8.0.0 ou posterior)

- [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

- A versão mais recente do [Yeoman](https://yeoman.io/) e do [Yeoman gerador de suplementos do Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Mesmo que você já tenha instalado o gerador Yeoman, recomendamos atualizar seu pacote para a versão mais recente do NPM.

## <a name="build-your-first-custom-functions-project"></a>Criar seu primeiro projeto de funções personalizadas

Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas. Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.

1. Execute o comando a seguir e responda aos prompts da seguinte forma.

    ```
    yo office
    ```

    - Escolha o tipo de projeto:`Excel Custom Functions Add-in project (...)`

    - Escolha um tipo de script: `JavaScript`

    - Qual será o nome do suplemento? `stock-ticker`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/12-10-fork-cf-pic.jpg)

    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.

2. Navegue até a pasta do projeto que você acabou de criar.

    ```
    cd stock-ticker
    ```

3. Confie no certificado autoassinado necessário para executar este projeto. Para obter instruções detalhadas para Windows ou Mac, confira [Adicionando Certificados Autoassinados como Certificado Raiz Confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Crie um projeto.

    ```
    npm run build
    ```

5. Inicie o servidor local da web, que é executado no Node.

    - Se você usar o Excel para Windows para testar suas funções personalizadas, execute o seguinte comando para iniciar o servidor Web local, inicie o Excel e Sideload o suplemento:

        ```
         npm run start
        ```
        Depois de executar esse comando, o prompt de comando mostrará detalhes sobre como iniciar o servidor Web. O Excel começará com seu suplemento carregado. Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente.

    - Se você usar o Excel online para testar suas funções personalizadas, execute o seguinte comando para iniciar o servidor Web local:

        ```
        npm run start-web
        ```

         Depois de executar esse comando, o prompt de comando mostrará detalhes sobre como iniciar o servidor Web. Para usar suas funções, abra uma nova pasta de trabalho no Excel online. Nesta pasta de trabalho, você precisará carregar o suplemento. 

        Para fazer isso, selecione a guia **Inserir** na faixa de opções e selecione **obter suplementos**. Na nova janela resultante, verifique se você está na guia **meus suplementos** . Em seguida, selecione **gerenciar meus suplementos _GT_ carregar meu suplemento**. Procure o arquivo de manifesto e carregue-o. Se o suplemento não for carregado, verifique se você concluiu a etapa 3 corretamente.

## <a name="try-out-the-prebuilt-custom-functions"></a>Experimentar as funções personalizadas predefinidas

O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas definidas no arquivo **src/customfunction.js**. O arquivo **manifest. XML** na pasta raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.

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
