---
title: Carregar suplementos do Office para Office na Web
description: Teste o suplemento do Office Office na Web por sideload.
ms.date: 09/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: 128e3537ac0ece5b7574dfec6d9d5c67b8d95a7b
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810377"
---
# <a name="sideload-office-add-ins-to-office-on-the-web"></a>Carregar suplementos do Office para Office na Web

Ao carregar um suplemento, você poderá instalar o suplemento sem colocá-lo primeiro em um catálogo de suplementos. Isso é útil ao testar e desenvolver seu suplemento, pois você pode ver como seu suplemento aparecerá e funcionará.

Quando você carrega um suplemento na Web, o manifesto do suplemento é armazenado no armazenamento local do navegador, portanto, se você limpar o cache do navegador ou mudar para um navegador diferente, precisará carregar o suplemento novamente.

As etapas para sideload de um suplemento na Web variam de acordo com os fatores a seguir.

- O aplicativo host (por exemplo, Excel, Word, Outlook)
- Qual ferramenta criou o projeto de suplemento (por exemplo, Visual Studio, gerador Yeoman para Suplementos do Office ou nenhum deles)
- Se você está com sideload para Office na Web com uma conta Microsoft ou com uma conta em um locatário do Microsoft 365

Na lista a seguir, vá para a seção ou artigo que corresponda ao seu cenário. Observe que o primeiro cenário na lista se aplica aos suplementos do Outlook. Os cenários restantes se aplicam a suplementos que não são do Outlook.

- Se você estiver carregando um suplemento do Outlook, consulte o artigo [Suplementos do Sideload outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md).
- Se você criou o suplemento usando o [gerador Yeoman para Suplementos do Office](../develop/yeoman-generator-overview.md), consulte [Sideload um suplemento criado pelo Yeoman para Office na Web](#sideload-a-yeoman-created-add-in-to-office-on-the-web).
- Se você criou o suplemento usando o Visual Studio, consulte [Sideload de um suplemento na Web ao usar o Visual Studio](#sideload-an-add-in-on-the-web-when-using-visual-studio).
- Para todos os outros casos, consulte uma das seções a seguir.

  - Se você estiver carregando de lado para Office na Web com uma conta Microsoft, consulte [Carregar manualmente um suplemento para Office na Web](#manually-sideload-an-add-in-to-office-on-the-web).
  - Se você estiver fazendo sideload para Office na Web com uma conta em um locatário do Microsoft 365, consulte [Sideload um suplemento para o Microsoft 365](#sideload-an-add-in-to-microsoft-365).

## <a name="sideload-a-yeoman-created-add-in-to-office-on-the-web"></a>Carregar um suplemento criado pelo Yeoman para Office na Web

Esse processo tem suporte apenas para **Excel**, **OneNote**, **PowerPoint** e **Word** . Este projeto de exemplo pressupõe que você esteja usando um projeto criado com o [gerador Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md).

1. Abra [Office na Web](https://office.live.com/) ou OneDrive. Usando a opção **Criar** , faça um documento no **Excel**, **OneNote**, **PowerPoint** ou **Word**. Neste novo documento, selecione **Compartilhar**, selecione **Copiar Link** e copie a URL.

1. Na linha de comando que começa no diretório raiz do seu projeto, execute o comando a seguir. Substitua "{url}" pela URL copiada.

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. Na primeira vez que você usar esse método para sideload de um suplemento na Web, você verá uma caixa de diálogo solicitando que você habilite o modo de desenvolvedor. Selecione a caixa de seleção **habilitar o Modo de Desenvolvedor agora** e selecione **OK**.

1. Você verá uma segunda caixa de diálogo, perguntando se deseja registrar um manifesto do Suplemento do Office no computador. Selecione **Sim**.

1. Seu suplemento está instalado. Se ele tiver um comando de suplemento, ele deverá aparecer na faixa de opções ou no menu de contexto. Se for um suplemento de painel de tarefas sem comandos de suplemento, o painel de tarefas deverá ser exibido.

## <a name="sideload-an-add-in-on-the-web-when-using-visual-studio"></a>Carregar um suplemento na Web ao usar o Visual Studio

Se você estiver usando o Visual Studio para desenvolver seu suplemento, pressione **F5** para abrir um documento do Office na *área de trabalho* , criar um documento em branco e carregar o suplemento. Quando você deseja sideload para *Office na Web*, o processo para sideload é semelhante ao sideload manual para a Web. A única diferença é que você deve atualizar o valor do elemento **SourceURL** e, possivelmente, outros elementos, em seu manifesto para incluir a URL completa em que o suplemento é implantado.

1. No Visual Studio, escolha **Exibir** > **Janela Propriedades**.

1. No **Gerenciador de Soluções**, selecione o projeto Web. Isso exibe propriedades para o projeto na janela **Propriedades** .

1. Na janela Propriedades, copie a **URL de SSL**.

1. No projeto de suplemento, abra o arquivo XML do manifesto. Certifique-se de que você está editando o XML de origem. Para alguns tipos de projeto, o Visual Studio abrirá uma exibição visual do XML que não funcionará para a próxima etapa.

1. Pesquisar e substituir todas as instâncias de **~remoteAppUrl/** pela URL de SSL que você copiou. Você verá várias substituições dependendo do tipo de projeto e as novas URLs serão semelhantes a `https://localhost:44300/Home.html`.

1. **Salve** o arquivo XML.

1. No **Gerenciador de Soluções**, abra o menu de contexto do projeto Web (por exemplo, clicando com o botão direito do mouse nele) e escolha **Depurar** > **Iniciar nova instância**. Isso executa o projeto Web sem iniciar o Office.

1. De Office na Web, faça sideload do suplemento usando as etapas descritas em [Carregar manualmente um suplemento para Office na Web](#manually-sideload-an-add-in-to-office-on-the-web).

## <a name="manually-sideload-an-add-in-to-office-on-the-web"></a>Carregar manualmente um suplemento para Office na Web

Esse método não usa a linha de comando e pode ser realizado usando comandos somente no aplicativo host (como o Excel).

1. Abra [Office na Web](https://office.com/). Abra um documento no **Excel**, **OneNote**, **PowerPoint** ou  **Word**. 

1. Na guia **Inserir** , na seção **Suplementos** , escolha **Suplementos do Office**.

1. Na caixa **de diálogo Suplementos do Office** , selecione a guia **MY ADD-INS** , escolha **Gerenciar Meus Suplementos** e, em seguida, **Carregar Meu Suplemento**.

    ![A caixa de diálogo Suplementos do Office com uma lista suspensa na leitura superior direita "Gerenciar meus suplementos" e uma lista suspensa abaixo dela com a opção "Carregar Meu Suplemento".](../images/office-add-ins-my-account.png)

1. **Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.

    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

1. Verifique se o suplemento está instalado. Por exemplo, se ele tiver um comando de suplemento, ele deverá aparecer na faixa de opções ou no menu de contexto. Se for um suplemento de painel de tarefas que não tem comandos de suplemento, o painel de tarefas deverá ser exibido.

> [!NOTE]
> Para testar seu Suplemento do Office com o Microsoft Edge com o WebView original (EdgeHTML), é necessária uma etapa de configuração adicional. Em um Prompt de Comando do Windows, execute a seguinte linha: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`. Isso não é necessário quando o Office está usando o Edge WebView2 baseado em Chromium. Para obter mais informações, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

[!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

## <a name="sideload-an-add-in-to-microsoft-365"></a>Carregar um suplemento no Microsoft 365

1. Entre em sua conta do Microsoft 365.

1. Abra o Iniciador de Aplicativos na extremidade esquerda da barra de ferramentas e selecione **Excel**, **OneNote**, **PowerPoint** ou **Word** e crie um novo documento.

1. Na guia **Inserir** , selecione o botão **Suplementos** .

1. Siga as etapas 3 a 5 da seção [Carregar manualmente um suplemento para Office na Web](#manually-sideload-an-add-in-to-office-on-the-web).

## <a name="remove-a-sideloaded-add-in"></a>Remover um suplemento sideload

Para remover um side-in carregado para Office na Web, basta limpar o cache do navegador. Se você fizer alterações no manifesto do suplemento (por exemplo, atualizar nomes de arquivo de ícones ou texto de comandos de suplemento), talvez seja necessário limpar o cache do navegador e recarregar o suplemento usando o manifesto atualizado. Isso permite que Office na Web renderize o suplemento conforme é descrito pelo manifesto atualizado.

## <a name="see-also"></a>Confira também

- [Sideload de suplementos do Office no Mac](sideload-an-office-add-in-on-mac.md)
- [Sideload de suplementos do Office no iPad](sideload-an-office-add-in-on-ipad.md)
- [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Limpar o cache do Office](clear-cache.md)
