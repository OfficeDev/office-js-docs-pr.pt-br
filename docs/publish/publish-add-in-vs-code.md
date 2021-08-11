---
title: Publicar um complemento usando o Visual Studio Code e o Azure
description: Como publicar um complemento usando Visual Studio Code e Azure Active Directory
ms.date: 08/12/2020
localization_priority: Normal
ms.openlocfilehash: e8c81a57b49254103366c28092f30235cc525e12d9a446897d862af4fc189325
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097204"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Publicar um suplemento desenvolvido com o Código do Visual Studio

Este artigo descreve como publicar um Suplemento do Office criado com o gerador Yeoman e desenvolvido com o [Código do Visual Studio (VS Code)](https://code.visualstudio.com) ou qualquer outro editor.

> [!NOTE]
> Para saber mais sobre como publicar um Suplemento do Office criado usando o Visual Studio, confira [Publicar o suplemento usando o Visual Studio](package-your-add-in-using-visual-studio.md).

## <a name="publishing-an-add-in-for-other-users-to-access"></a>Publicar um suplemento para que outros usuários acessem o

Um Suplemento do Office consiste em um aplicativo Web e um arquivo de manifesto. O aplicativo Web define a interface do usuário e a funcionalidade do suplemento, enquanto o manifesto especifica o local do aplicativo Web e define as configurações e os recursos do suplemento.

Enquanto estiver desenvolvendo, você pode executar o complemento em seu servidor Web local ( `localhost` ). Quando estiver pronto para publicá-lo para outros usuários acessarem, você precisará implantar o aplicativo Web e atualizar o manifesto para especificar a URL do aplicativo implantado.

Quando o seu add-in estiver funcionando conforme desejado, você poderá publicá-lo diretamente Visual Studio Code usando a extensão do Azure Armazenamento.

## <a name="using-visual-studio-code-to-publish"></a>Usando Visual Studio Code para publicar

>[!NOTE]
> Essas etapas só funcionam para projetos criados com o gerador Yeoman.

1. Abra seu projeto de sua pasta raiz em Visual Studio Code (VS Code).
2. Na exibição Extensões em VS Code, pesquise a extensão do Azure Armazenamento e instale-a.
3. Depois de instalado, um ícone do Azure é adicionado à Barra de Atividades. Selecione-o para acessar a extensão. Se sua Barra de Atividades estiver oculta, você não poderá acessar a extensão. Mostrar a Barra de Atividades **selecionando Exibir > Aparência > Mostrar Barra de Atividades**.
4. Quando estiver na extensão, entre em sua conta do Azure selecionando **Entrar no Azure**. Você também pode criar uma conta do Azure se ainda não tiver uma selecionando **Criar uma conta gratuita do Azure**. Siga as etapas fornecidas para configurar sua conta.
5. Depois de entrar na sua conta do Azure, você verá suas contas de armazenamento do Azure aparecerem na extensão. Se você ainda não tiver uma conta de armazenamento, precisará criar uma usando a **opção Criar nova** conta de armazenamento. Nomeia sua conta de armazenamento como um nome global exclusivo, usando apenas 'a-z' e '0-9'. Observe que, por padrão, isso cria uma conta de armazenamento e um grupo de recursos com o mesmo nome. Ele coloca automaticamente a conta de armazenamento no Oeste dos EUA. Isso pode ser ajustado online por [meio de sua conta do Azure.](https://portal.azure.com/)
6. Selecione e segure (clique com o botão direito do mouse) em sua conta de armazenamento, escolhendo **Configurar site estático.** Você será solicitado a inserir o nome do documento de índice e o nome do documento 404. Altere o nome do documento de índice do padrão `index.html` para **`taskpane.html`** . Você também pode optar por alterar o nome do documento 404, mas não é necessário.
7. Selecione e segure (clique com o botão direito do mouse) no armazenamento novamente, desta vez escolhendo **Procurar site estático**. Na janela do navegador aberta, copie a URL do site.
8. Em VS Code, abra o arquivo de manifesto do projeto ( ) e altere qualquer referência à URL do seu localhost (como ) para a `manifest.xml` `https://localhost:3000` URL que você copiou. Esse ponto de extremidade é a URL do site estático para sua conta de armazenamento recém-criada. Salve as alterações no arquivo de manifesto.
9. Abra um prompt de linha de comando e navegue até o diretório raiz do seu projeto de complemento. Em seguida, execute o seguinte comando para preparar todos os arquivos para implantação de produção.

    ```command&nbsp;line
    npm run build
    ```

    Quando a compilação for concluída, a pasta **dist** no diretório raiz do projeto de suplemento incluirá os arquivos que você implantará nas etapas subsequentes.

10. Para implantar, selecione o explorador de arquivos, selecione e segure (clique com o botão direito do mouse) em sua **pasta dist** e escolha **Implantar no Site Estático**. Quando solicitado, selecione a conta de armazenamento criada anteriormente.

![Implantando em um site estático.](../images/deploy-to-static-website.png)

11. Quando a implantação é concluída, uma **mensagem Procurar** para site é exibida que você pode selecionar para abrir o ponto de extremidade principal do código do aplicativo implantado.

## <a name="see-also"></a>Confira também

- [Desenvolver Suplementos do Office com o Código do Visual Studio](../develop/develop-add-ins-vscode.md)
- [Implantar e publicar seu suplemento do Office](../publish/publish.md)
