---
title: Publicar um suplemento usando o Visual Studio Code e o Azure
description: Como publicar um suplemento usando o Visual Studio Code e o Azure Active Directory
ms.date: 08/12/2020
localization_priority: Normal
ms.openlocfilehash: 3552e4eebacc84fc2b8e37782c97b4e03e96e508
ms.sourcegitcommit: 7faa0932b953a4983a80af70f49d116c3236d81a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/21/2020
ms.locfileid: "46845505"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Publicar um suplemento desenvolvido com o Código do Visual Studio

Este artigo descreve como publicar um Suplemento do Office criado com o gerador Yeoman e desenvolvido com o [Código do Visual Studio (VS Code)](https://code.visualstudio.com) ou qualquer outro editor.

> [!NOTE]
> Para saber mais sobre como publicar um Suplemento do Office criado usando o Visual Studio, confira [Publicar o suplemento usando o Visual Studio](package-your-add-in-using-visual-studio.md).

## <a name="publishing-an-add-in-for-other-users-to-access"></a>Publicar um suplemento para que outros usuários acessem o

Um Suplemento do Office é formado por um aplicativo Web e um arquivo de manifesto. O aplicativo Web define a interface do usuário e a funcionalidade do suplemento, enquanto o manifesto especifica o local do aplicativo Web e define as configurações e os recursos do suplemento.

Enquanto estiver desenvolvendo, você poderá executar o suplemento no servidor Web local ( `localhost` ). Quando estiver pronto para publicá-lo para que outros usuários o acessem, você precisará implantar o aplicativo Web e atualizar o manifesto para especificar a URL do aplicativo implantado.

Quando o suplemento estiver funcionando conforme desejado, você poderá publicá-lo diretamente pelo código do Visual Studio usando a extensão de armazenamento do Azure.

## <a name="using-visual-studio-code-to-publish"></a>Usando o Visual Studio Code para publicar

>[!NOTE]
> Estas etapas só funcionam para projetos criados com o gerador Yeoman.

1. Abra o projeto na pasta raiz do Visual Studio Code (VS Code).
2. No modo de exibição de extensões no VS Code, procure a extensão de armazenamento do Azure e instale-a.
3. Depois de instalado, um ícone do Azure é adicionado à barra de atividade. Selecione-o para acessar a extensão. Se sua barra de atividades estiver oculta, você não poderá acessar a extensão. Mostrar a barra de atividades selecionando **exibir > aparência > Mostrar barra de atividades**.
4. Na extensão, entre em sua conta do Azure selecionando **entrar no Azure**. Você também pode criar uma conta do Azure se ainda não tiver um selecionando **criar uma conta gratuita do Azure**. Siga as etapas fornecidas para configurar sua conta.
5. Depois de entrar em sua conta do Azure, você verá suas contas de armazenamento do Azure exibidas na extensão. Se você ainda não tem uma conta de armazenamento, precisará criar uma usando a opção **criar nova conta de armazenamento** . Nomeie sua conta de armazenamento com um nome globalmente exclusivo, usando apenas ' a-z ' e ' 0-9 '. Observe que, por padrão, isso cria uma conta de armazenamento e um grupo de recursos com o mesmo nome. Ele coloca automaticamente a conta de armazenamento no oeste dos EUA. Isso pode ser ajustado online por meio [da sua conta do Azure](https://portal.azure.com/).
6. Selecione e segure (clique com o botão direito do mouse) sua conta de armazenamento, escolha **configurar site estático**. Você será solicitado a inserir o nome do documento de índice e o nome do documento 404. Altere o nome do documento de índice de padrão `index.html` para **`taskpane.html`** . Você pode decidir também alterar o nome do documento 404, mas não é necessário.
7. Selecione e segure (clique com o botão direito do mouse) seu armazenamento novamente, desta vez escolhendo **navegar no site estático**. Na janela do navegador que é aberta, copie a URL do site.
8. No VS Code, abra o arquivo de manifesto do seu projeto ( `manifest.xml` ) e altere qualquer referência à URL do localhost (como `https://localhost:3000` ) para a URL que você copiou. Este ponto de extremidade é a URL estática do site para sua conta de armazenamento recém-criada. Salve as alterações no arquivo de manifesto.
9. Abra um prompt de linha de comando e navegue até o diretório raiz do seu projeto de suplemento. Em seguida, execute o seguinte comando para preparar todos os arquivos para implantação de produção.

    ```command&nbsp;line
    npm run build
    ```

    Quando a compilação for concluída, a pasta **dist**no diretório raiz do projeto de suplemento incluirá os arquivos que você implantará nas etapas subsequentes.

10. Para implantar o, selecione o explorador de arquivos, selecione e segure (clique com o botão direito do mouse) sua pasta **dist** e escolha **implantar no site estático**. Quando solicitado, selecione a conta de armazenamento criada anteriormente.

![Implantando em um site estático](../images/deploy-to-static-website.png)

11. Quando a implantação estiver concluída, será exibida uma mensagem de **navegar para o site** que você pode selecionar para abrir o ponto de extremidade principal do código do aplicativo implantado.

## <a name="see-also"></a>Confira também

- [Desenvolver Suplementos do Office com o Código do Visual Studio](../develop/develop-add-ins-vscode.md)
- [Implantar e publicar seu suplemento do Office](../publish/publish.md)
