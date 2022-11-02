---
title: Publicar um suplemento usando Visual Studio Code e o Azure
description: Como publicar um suplemento usando Visual Studio Code e o Azure Active Directory
ms.date: 09/07/2022
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
ms.openlocfilehash: b2d05ba9fb1c20529731312dab112abe6a00cfc7
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810068"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Publicar um suplemento desenvolvido com o Código do Visual Studio

Este artigo descreve como publicar um Suplemento do Office criado com o gerador Yeoman e desenvolvido com o [Código do Visual Studio (VS Code)](https://code.visualstudio.com) ou qualquer outro editor.

> [!NOTE]
> Para saber mais sobre como publicar um Suplemento do Office criado usando o Visual Studio, confira [Publicar o suplemento usando o Visual Studio](package-your-add-in-using-visual-studio.md).

## <a name="publishing-an-add-in-for-other-users-to-access"></a>Publicar um suplemento para que outros usuários acessem o

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

Enquanto estiver desenvolvendo, você pode executar o suplemento no servidor Web local (`localhost`). Quando estiver pronto para publicá-lo para outros usuários acessarem, você precisará implantar o aplicativo Web e atualizar o manifesto para especificar a URL do aplicativo implantado.

Quando o suplemento estiver funcionando conforme desejado, você pode publicá-lo diretamente por meio de Visual Studio Code usando a extensão de Armazenamento do Azure.

## <a name="using-visual-studio-code-to-publish"></a>Usando o Visual Studio Code para publicar

>[!NOTE]
> Essas etapas funcionam apenas para projetos criados com o gerador Yeoman.

1. Abra seu projeto de sua pasta raiz em Visual Studio Code (VS Code).
1. Selecione **Exibir** > **Extensões** (Ctrl+Shift+X) para abrir a exibição Extensões.
1. Pesquise a extensão **de Armazenamento do Azure** e instale-a.
1. Depois de instalado, um ícone do Azure será adicionado à **Barra de Atividades**. Selecione-a para acessar a extensão. Se a **Barra de Atividades** estiver oculta, abra-a selecionando **Exibir** > **Barra de Atividades** de **Aparência** > .
1. Selecione **Entrar no Azure** para entrar em sua conta do Azure. Se você ainda não tiver uma conta do Azure, crie uma selecionando **Criar uma Conta do Azure**. Siga as etapas fornecidas para configurar sua conta.

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="Entre no botão do Azure selecionado na extensão do Azure.":::

1. Depois de entrar, você verá suas contas de armazenamento do Azure aparecerem na extensão. Se você ainda não tiver uma conta de armazenamento, crie uma usando a opção **Criar Conta de Armazenamento** na paleta de comandos. Nomeie sua conta de armazenamento como um nome globalmente exclusivo, usando apenas 'a-z' e '0-9'. Observe que, por padrão, isso cria uma conta de armazenamento e um grupo de recursos com o mesmo nome. Ele coloca automaticamente a conta de armazenamento no oeste dos EUA. Isso pode ser ajustado online por meio de [sua conta do Azure](https://portal.azure.com/).

    :::image type="content" source="../images/azure-extension-create-storage-account.png" alt-text="Selecionar contas de armazenamento > Criar Conta de Armazenamento na extensão do Azure.":::

1. Clique com o botão direito do mouse em sua conta de armazenamento e **selecione Configurar Site Estático**. Você será solicitado a inserir o nome do documento de índice e o nome do documento 404. Altere o nome do documento de índice do padrão `index.html` para **`taskpane.html`**. Você também pode alterar o nome do documento 404, mas não é necessário.
1. Clique com o botão direito do mouse em sua conta de armazenamento novamente e, desta vez, **selecione Procurar Site Estático**. Na janela do navegador aberta, copie a URL do site.
1. Abra o arquivo de manifesto do projeto (`manifest.xml`) e altere todas as referências à URL de localhost (como `https://localhost:3000`) para a URL copiada. Esse ponto de extremidade é a URL do site estático para sua conta de armazenamento recém-criada. Salve as alterações no arquivo de manifesto.
1. Abra um prompt de linha de comando ou uma janela de terminal e vá para o diretório raiz do seu projeto de suplemento. Execute o comando a seguir para preparar todos os arquivos para implantação de produção.

    ```command&nbsp;line
    npm run build
    ```

    Quando a compilação for concluída, a pasta **dist** no diretório raiz do projeto de suplemento incluirá os arquivos que você implantará nas etapas subsequentes.

1. No VS Code, acesse o Explorer e clique com o botão direito do mouse na pasta **dist** e **selecione Implantar no Site Estático por meio do Armazenamento do Azure**. Quando solicitado, selecione a conta de armazenamento que você criou anteriormente.

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="Selecione a pasta dist, clique com o botão direito do mouse e selecione Implantar no Site Estático por meio do Armazenamento do Azure.":::

1. Quando a implantação for concluída, clique com o botão direito do mouse na conta de armazenamento que você criou anteriormente e selecione **Procurar Site Estático**. Isso abre o site estático e exibe o painel de tarefas.

1. Por fim, [o sideload do arquivo de manifesto](../testing/sideload-office-add-ins-for-testing.md) e o suplemento serão carregados do site estático que você acabou de implantar.

## <a name="deploy-custom-functions-for-excel"></a>Implantar funções personalizadas para o Excel

Se o suplemento tiver funções personalizadas, haverá mais algumas etapas para habilitá-las na conta de Armazenamento do Azure. Primeiro, habilite o CORS para que o Office possa acessar o arquivo functions.json.

1. Clique com o botão direito do mouse na conta de armazenamento do Azure e selecione **Abrir no Portal**.
1. No grupo Configurações, selecione **CORS (compartilhamento de recursos)**. Você também pode usar a caixa de pesquisa para encontrar isso.
1. Crie uma nova regra CORS com as seguintes configurações.

    |Propriedade        |Valor                        |
    |----------------|-----------------------------|
    |Origens permitidas | \*                          |
    |Métodos permitidos | OBTER                         |
    |Cabeçalhos permitidos | \*                          |
    |Cabeçalhos expostos | Access-Control-Allow-Origin |
    |Idade máxima         | 200                         |

1. Selecione **Salvar**.

> [!CAUTION]
> Essa configuração cors pressupõe que todos os arquivos em seu servidor estejam publicamente disponíveis para todos os domínios.  

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
1. Adicione o código a seguir na lista de `plugins` para copiar o web.config no pacote quando o build for executado.

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

    Quando o build for concluído, a pasta **dist** no diretório raiz do projeto de suplemento conterá os arquivos que você implantará.

1. Para implantar, no VS Code **Explorer**, clique com o botão direito do mouse na pasta **dist** e **selecione Implantar no Site Estático por meio do Armazenamento do Azure**. Quando solicitado, selecione a conta de armazenamento que você criou anteriormente. Se você já implantou a pasta **dist** , será solicitado se quiser substituir os arquivos no armazenamento do Azure com as alterações mais recentes.

## <a name="see-also"></a>Confira também

- [Desenvolver Suplementos do Office com o Código do Visual Studio](../develop/develop-add-ins-vscode.md)
- [Implantar e publicar seu suplemento do Office](../publish/publish.md)
- [Suporte ao CORS (Compartilhamento de Recursos de Origem Cruzada) para Armazenamento do Azure](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)
