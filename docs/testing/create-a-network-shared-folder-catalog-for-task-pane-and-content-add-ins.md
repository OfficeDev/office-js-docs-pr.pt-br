---
title: Realizar sideload de Suplementos do Office para teste de um compartilhamento de rede
description: Saiba como fazer sideload de um Suplemento do Office para teste de um compartilhamento de rede.
ms.date: 05/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e32e91a542ec4c3557ff945b93d69fc12eac766c
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659833"
---
# <a name="sideload-office-add-ins-for-testing-from-a-network-share"></a>Realizar sideload de Suplementos do Office para teste de um compartilhamento de rede

Você pode testar um Suplemento do Office em um cliente do Office que está no Windows publicando o manifesto em um compartilhamento de arquivos de rede (instruções abaixo). Essa opção de implantação destina-se a ser usada quando você tiver concluído o desenvolvimento e o teste em um localhost e quiser testar o suplemento de uma conta de nuvem ou servidor não local.

> [!IMPORTANT]
> A implantação por compartilhamento de rede não tem suporte para suplementos de produção. Esse método tem as seguintes limitações.
>
> - O suplemento só pode ser instalado em computadores Windows.
> - Se uma nova versão de um suplemento alterar a faixa de opções, como adicionar uma guia personalizada ou um botão personalizado a ela, cada usuário precisará reinstalar o suplemento.

> [!NOTE]
> Se o projeto de suplemento tiver sido criado com uma versão suficientemente recente do [Gerador Yeoman para Suplementos do Office](../develop/yeoman-generator-overview.md), o suplemento realizará sideload automaticamente no cliente de desktop do Office ao executar o `npm start`.

Este artigo se aplica somente ao teste de suplementos do Word, Excel, PowerPoint e Project e somente no Windows. Se você quiser testar em outra plataforma ou quiser testar um suplemento do Outlook, consulte um dos tópicos a seguir para realizar o sideload do suplemento.

- [Realizar sideload de suplementos do Office no Office na Web para teste](sideload-office-add-ins-for-testing.md)
- [Sideload suplementos do Office para teste em um iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)

O vídeo a seguir oferece orientações para a realização do processo de sideload no suplemento do Office na Web ou para área de trabalho usando um catálogo de pasta compartilhada.  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a>Compartilhar uma pasta

1. No computador do Windows, onde você deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que você deseja usar como seu catálogo de pasta compartilhada.

1. Abra o menu de contexto na pasta que você deseja usar como catálogo de pasta compartilhada (clique com o botão direito) e escolha **Propriedades**.

1. Dentro da janela de diálogo **Propriedades** abra a guia **Compartilhamento** e escolha o botão **Compartilhar**.

    ![Caixa de diálogo Propriedades da Pasta com a guia Compartilhamento e o botão Compartilhar realçados.](../images/sideload-windows-properties-dialog.png)

1. Dentro a janela de diálogo **Acesso à rede** adicione você mesmo e quaisquer outros usuários e/ou grupos com quem você deseja compartilhar o suplemento. Você precisará de pelo menos da permissão **Leitura/Gravação** para a pasta. Quando terminar de escolher as pessoas para compartilhar, escolha o botão **Compartilhar**.

1. Quando você vir a confirmação **Sua pasta foi compartilhada**, anote o caminho de rede completo que é exibido imediatamente após o nome da pasta. (Você precisará inserir esse valor como o **Url Catálogo** quando você [especificar a pasta compartilhada como um catálogo confiável](#specify-the-shared-folder-as-a-trusted-catalog), conforme descrito na próxima seção deste artigo.) Escolha o botão **Concluído** para fechar a janela de diálogo **Acesso à rede**.

   ![Caixa de diálogo de acesso à rede com o caminho de compartilhamento realçado.](../images/sideload-windows-network-access-dialog.png)

1. Escolha o botão **Fechar** para fechar a caixa de diálogo **Propriedades**.

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>Especifique a pasta compartilhada como um catálogo confiável

### <a name="configure-the-trust-manually"></a>Configure a confiança manualmente

1. Abra um novo documento no Excel, no Word, no PowerPoint ou no Project.

1. Escolha a guia **Arquivo** e, então, **Opções**.

1. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.

1. Escolha **Catálogos de Suplemento Confiáveis**.

1. Na caixa **Url catálogo**, digite o caminho completo da rede para a pasta que você [compartilhou](#share-a-folder) anteriormente. Se você não conseguiu anotar todo o caminho de rede da pasta ao compartilhar a pasta, você pode obtê-lo na janela de diálogo **Propriedades**, conforme mostrado na captura de tela a seguir.

    ![Caixa de diálogo Propriedades da Pasta com a guia Compartilhamento e o caminho de rede realçados.](../images/sideload-windows-properties-dialog-2.png)

1. Depois de inserir o caminho de de rede completo da pasta na caixa **Url catálogo**, escolha o botão **Adicionar Catálogo**.

1. Selecione a caixa de seleção **Mostrar no Menu** no novo item adicionado e, em seguida, escolha o botão **Ok** para fechar a janela de diálogo **Central de Confiabilidade**. 

    ![Caixa de diálogo Central de Confiabilidade com o catálogo selecionado.](../images/sideload-windows-trust-center-dialog.png)

1. Escolha o **botão OK** para fechar a janela **de** diálogo Opções.

1. Feche e abra novamente o aplicativo do Office para que as alterações tenham efeito.

### <a name="configure-the-trust-with-a-registry-script"></a>Configurar a confiança com um script de Registro

1. Em um editor de texto, crie um arquivo chamado TrustNetworkShareCatalog.reg.

1. Adicione o conteúdo a seguir ao arquivo.

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```

1. Use uma das várias ferramentas de geração de GUID online, como o [Gerador de GUIDs](https://guidgenerator.com/), para gerar um GUID aleatório e, no arquivo TrustNetworkShareCatalog.reg, substitua a cadeia de caracteres "-random-GUID-here-" *nos dois locais* pelo GUID. (Os símbolos `{}` de delimitação devem permanecer.)

1. Substitua o valor `Url` pelo caminho completo da rede para a pasta que você [compartilhou](#share-a-folder) anteriormente. (Observe que quaisquer caracteres `\` na URL devem ser duplicados.) Se você não conseguiu anotar todo o caminho de rede da pasta ao compartilhar a pasta, você pode obtê-lo na janela de diálogo **Propriedades**, conforme mostrado na captura de tela a seguir.

    ![Caixa de diálogo Propriedades da Pasta com a guia Compartilhamento e o caminho de rede realçados.](../images/sideload-windows-properties-dialog-2.png)

1. Agora o arquivo deve ter a aparência a seguir. Salve-o.

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

1. Feche *todos* os aplicativos do Office.

1. Execute o TrustNetworkShareCatalog.reg como faria com qualquer arquivo executável, por exemplo, com um clique duplo.

## <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento

1. Coloque o arquivo de manifesto XML de qualquer suplemento que você esteja testando no catálogo de pasta compartilhada. Observe que você implanta o próprio aplicativo Web em um servidor Web. Especifique a URL no elemento **\<SourceLocation\>** do arquivo de manifesto.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

    > [!NOTE]
    > Para projetos do Visual Studio, use o manifesto criado pelo projeto na `{projectfolder}\bin\Debug\OfficeAppManifests` pasta.

1. No Excel, Word ou PowerPoint, selecione **Meus Suplementos** na guia **Inserir** da faixa de opções. No Project, selecione **Meus Suplementos** na guia **Projeto** da faixa de opções.

1. Escolha **PASTA COMPARTILHADA** na parte superior da caixa de diálogo **Suplementos do Office**.

1. Selecione o nome do suplemento e escolha **Adicionar** para inseri-lo.

## <a name="remove-a-sideloaded-add-in"></a>Remover um suplemento de sideload

Você pode remover um suplemento com sideload anteriormente limpando o cache do Office em seu computador. Detalhes sobre como limpar o cache no Windows podem ser encontrados no artigo [Limpar o cache do Office](clear-cache.md#clear-the-office-cache-on-windows).

## <a name="see-also"></a>Confira também

- [Validar o manifesto de Suplemento do Office](troubleshoot-manifest.md)
- [Limpar o cache do Office](clear-cache.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)
