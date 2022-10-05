---
title: Publicar um suplemento usando o Visual Studio Code e o Azure
description: Como publicar um suplemento usando o Visual Studio Code e o Azure Active Directory
ms.date: 09/07/2022
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
ms.openlocfilehash: f1e931042f7796333bbc75b2666f3564e63b9f7b
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467185"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Publicar um suplemento desenvolvido com o Código do Visual Studio

Este artigo descreve como publicar um Suplemento do Office criado com o gerador Yeoman e desenvolvido com o [Código do Visual Studio (VS Code)](https://code.visualstudio.com) ou qualquer outro editor.

> [!NOTE]
> Para saber mais sobre como publicar um Suplemento do Office criado usando o Visual Studio, confira [Publicar o suplemento usando o Visual Studio](package-your-add-in-using-visual-studio.md).

## <a name="publishing-an-add-in-for-other-users-to-access"></a>Publicar um suplemento para que outros usuários acessem o

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

Durante o desenvolvimento, você pode executar o suplemento no servidor Web local (`localhost`). Quando estiver pronto para publicá-lo para outros usuários acessarem, você precisará implantar o aplicativo Web e atualizar o manifesto para especificar a URL do aplicativo implantado.

Quando o suplemento estiver funcionando conforme desejado, você poderá publicá-lo diretamente por meio Visual Studio Code usando a extensão de Armazenamento do Azure.

## <a name="using-visual-studio-code-to-publish"></a>Usando o Visual Studio Code para publicar

>[!NOTE]
> Essas etapas só funcionam para projetos criados com o gerador Yeoman.

1. Abra seu projeto na pasta raiz no Visual Studio Code (VS Code).
1. Selecione **Exibir > Extensões** (Ctrl+Shift+X) para abrir o modo de exibição Extensões.
1. Pesquise a **extensão de Armazenamento do Azure** e instale-a.
1. Depois de instalado, um ícone do Azure é adicionado à Barra **de Atividades**. Selecione-o para acessar a extensão. Se a **Barra de Atividades** estiver oculta, abra-a selecionando **Exibir > Aparência > Barra de Atividades**.
1. Selecione **Entrar no Azure** para entrar em sua conta do Azure. Se você ainda não tiver uma conta do Azure, crie uma selecionando **Criar uma Conta do Azure**. Siga as etapas fornecidas para configurar sua conta.

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="Botão Entrar no Azure selecionado na extensão do Azure.":::

1. Depois de entrar, você verá suas contas de armazenamento do Azure aparecerem na extensão. Se você ainda não tiver uma conta de armazenamento, crie uma usando a **opção** Criar Conta de Armazenamento na paleta de comandos. Nomeie sua conta de armazenamento como um nome globalmente exclusivo, usando apenas 'a-z' e '0-9'. Observe que, por padrão, isso cria uma conta de armazenamento e um grupo de recursos com o mesmo nome. Ele coloca automaticamente a conta de armazenamento no Oeste dos EUA. Isso pode ser ajustado online por [meio de sua conta do Azure](https://portal.azure.com/).

    :::image type="content" source="../images/azure-extension-create-storage-account.png" alt-text="Selecionar contas de armazenamento > criar conta de armazenamento na extensão do Azure.":::

1. Clique com o botão direito do mouse em sua conta de armazenamento e selecione **Configurar Site Estático**. Será solicitado que você insira o nome do documento de índice e o nome do documento 404. Altere o nome do documento de índice do padrão `index.html` para **`taskpane.html`**. Você também pode alterar o nome do documento 404, mas não é necessário.
1. Clique com o botão direito do mouse em sua conta de armazenamento novamente e, desta vez, selecione **Procurar Site Estático**. Na janela do navegador que é aberta, copie a URL do site.
1. Abra o arquivo de manifesto do projeto (`manifest.xml`) e altere todas as referências à URL do localhost ( `https://localhost:3000`como) para a URL que você copiou. Esse ponto de extremidade é a URL do site estático para sua conta de armazenamento recém-criada. Salve as alterações no arquivo de manifesto.
1. Abra um prompt de linha de comando ou uma janela de terminal e vá para o diretório raiz do seu projeto de suplemento. Execute o comando a seguir para preparar todos os arquivos para implantação de produção.

    ```command&nbsp;line
    npm run build
    ```

    Quando a compilação for concluída, a pasta **dist** no diretório raiz do projeto de suplemento incluirá os arquivos que você implantará nas etapas subsequentes.

1. No VS Code, vá para o Explorer e clique com o botão direito do mouse na pasta **dist** e selecione Implantar no Site Estático **por meio do Armazenamento do Azure**. Quando solicitado, selecione a conta de armazenamento criada anteriormente.

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="Selecione a pasta dist, clique com o botão direito do mouse e selecione Implantar no Site Estático por meio do Armazenamento do Azure.":::

1. Quando a implantação for concluída, clique com o botão direito do mouse na conta de armazenamento que você criou anteriormente e selecione **Procurar Site Estático**. Isso abre o site estático e exibe o painel de tarefas.

1. Por fim, [o sideload do arquivo de manifesto](../testing/sideload-office-add-ins-for-testing.md) e o suplemento será carregado do site estático que você acabou de implantar.

## <a name="deploy-custom-functions-for-excel"></a>Implantar funções personalizadas para o Excel

Se o suplemento tiver funções personalizadas, haverá mais algumas etapas para habilita-las na conta de Armazenamento do Azure. Primeiro, habilite o CORS para que o Office possa acessar o arquivo functions.json.

1. Clique com o botão direito do mouse na conta de armazenamento do Azure e selecione **Abrir no Portal**.
1. No grupo Configurações, selecione **Compartilhamento de recursos (CORS)**. Você também pode usar a caixa de pesquisa para encontrar isso.
1. Crie uma nova regra CORS com as configurações a seguir.

    |Propriedade        |Valor                        |
    |----------------|-----------------------------|
    |Origens permitidas | \*                          |
    |Métodos permitidos | OBTER                         |
    |Cabeçalhos permitidos | \*                          |
    |Cabeçalhos expostos | Access-Control-Allow-Origin |
    |Idade máxima         | 200                         |

1. Selecione **Salvar**.

> [!CAUTION]
> Essa configuração do CORS pressupõe que todos os arquivos em seu servidor estejam disponíveis publicamente para todos os domínios.  

Em seguida, adicione um tipo MIME para arquivos JSON.

1. Crie um novo arquivo na pasta /src chamada **web.config**.
1. Insira o XML a seguir e salve o arquivo.

    ```xml
    <?xml version="1.0"?>
    <configuration>
      <system.webServer>
        <staticContent>
          <mimeMap fileExtension=".json" mimeType="application/json" />
        </staticContent>
      </system.webServer>
    </configuration> 
    ```

1. Abra o arquivo **webpack.config.js**.
1. Adicione o código a seguir na lista para `plugins` copiar o web.config no pacote quando o build for executado.

    ```javascript
    new CopyWebpackPlugin({
      patterns: [
      {
        from: "src/web.config",
        to: "src/web.config",
      },
     ],
    }),
    ```

1. Abra um prompt de linha de comando e vá para o diretório raiz do seu projeto de suplemento. Em seguida, execute o comando a seguir para preparar todos os arquivos para implantação.

    ```command&nbsp;line
    npm run build
    ```

    Quando o build for concluído, a pasta **dist** no diretório raiz do seu projeto de suplemento conterá os arquivos que você implantará.

1. Para implantar, no Vs Code **Explorer**, clique com o botão direito do mouse na pasta **dist** e selecione Implantar no Site Estático **por meio do Armazenamento do Azure**. Quando solicitado, selecione a conta de armazenamento criada anteriormente. Se você já implantou a pasta **dist** , será solicitado que você queira substituir os arquivos no armazenamento do Azure com as alterações mais recentes.

## <a name="see-also"></a>Confira também

- [Desenvolver Suplementos do Office com o Código do Visual Studio](../develop/develop-add-ins-vscode.md)
- [Implantar e publicar seu suplemento do Office](../publish/publish.md)
- [Suporte ao CORS (Compartilhamento de Recursos entre Origens) para o Armazenamento do Azure](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)
