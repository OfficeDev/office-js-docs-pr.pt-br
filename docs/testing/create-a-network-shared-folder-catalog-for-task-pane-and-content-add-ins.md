---
title: Fazer sideload de suplementos do Office para teste
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 6ee8e4e9a2413b34cb8991b09d61e16888a0e6a6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640019"
---
# <a name="sideload-office-add-ins-for-testing"></a>Fazer sideload de suplementos do Office para teste

Você pode instalar um suplemento do Office para teste em um cliente do Office no Windows publicando o manifesto em um compartilhamento de arquivos na rede (instruções abaixo).

> [!NOTE]
> Se seu projeto de suplemento foi criado com a [ferramenta **yo office**](https://github.com/OfficeDev/generator-office), há uma maneira alternativa de fazer sideload que pode servir para você. Para obter mais detalhes, consulte [Fazer sideload de suplementos do Office usando o comando de sideload](sideload-office-addin-using-sideload-command.md).

Este artigo se aplica somente para testar suplementos do Word, PowerPoint ou Excel no Windows. Se você deseja testar em outra plataforma ou deseja testar um suplemento do Outlook, consulte um dos seguintes tópicos para fazer sideload de seu suplemento:

- [Fazer sideload de suplementos do Office para teste no Office Online](sideload-office-add-ins-for-testing.md)
- [Sideload dos suplementos do Office para teste em um iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Fazer sideload de suplementos do Outlook para teste](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


O vídeo a seguir oferece orientações para o processo de sideload do seu suplemento na área de trabalho do Office ou no Office Online usando um catálogo de pasta compartilhada.  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a>Compartilhar uma pasta

1. No Explorador de Arquivos no computador do Windows em que você deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que você deseja usar como seu catálogo de pasta compartilhada.

2. Abra o menu de contexto para a pasta que você deseja usar como seu catálogo de pasta compartilhada (clique com o botão direito do mouse na pasta) e escolha **Propriedades**.

3. Dentro da janela de diálogo **Propriedades** , abra a guia **Compartilhamento** e escolha o botão **Compartilhar** .

    ![caixa de diálogo Propriedades da pasta com a guia Compartilhamento e o botão Compartilhar realçados](../images/sideload-windows-properties-dialog.png)

4. Dentro da janela de diálogo **Acesso à rede**, adicione a si mesmo e quaisquer outros usuários e/ou grupos com quem você deseja compartilhar seu suplemento. Você precisará, no mínimo, de permissão de **Leitura/Gravação** para a pasta. Depois de concluir a seleção de pessoas com as quais fazer o compartilhamento, escolha o botão **Compartilhar**.

5. Quando você vir a confirmação de que **Sua pasta está compartilhada**, anote o caminho completo de rede que é exibido imediatamente após o nome da pasta. (Você precisará digitar esse valor como a **Url do Catálogo** quando você [especificar a pasta compartilhada como um catálogo confiável](#specify-the-shared-folder-as-a-trusted-catalog), conforme descrito na próxima seção deste artigo.) Escolha o botão **Concluído** para fechar a janela de diálogo de **Acesso à rede**.

   ![Caixa de diálogo de acesso à rede com o caminho de compartilhamento realçado](../images/sideload-windows-network-access-dialog.png)

6. Escolha o botão **Fechar** para fechar a janela de diálogo **Propriedades** .

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>Especificar a pasta compartilhada como um catálogo confiável
      
1. Abra um novo documento no Excel, no Word ou no PowerPoint.
    
2. Escolha a guia **Arquivo** e escolha **Opções**.
    
3. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.
    
4. Escolha **Catálogos de Suplemento Confiáveis**.
    
5. Na caixa **Url do Catálogo** , insira o caminho completo de rede para a pasta que você [compartilhou](#share-a-folder) anteriormente. Se você não conseguiu anotar o caminho de rede completo quando você compartilhou a pasta, você pode obtê-lo da janela de diálogo **Propriedades** da pasta, conforme mostrado na seguinte captura de tela. 

    ![diálogo Propriedades da pasta com a guia Compartilhamento e o caminho de rede realçados](../images/sideload-windows-properties-dialog-2.png)
    
6. Depois de inserir o caminho de rede completo da pasta na caixa **Url do Catálogo**, escolha o botão **Adicionar catálogo**.

7. Selecione a caixa de seleção **Mostrar no Menu** referente ao item recém-adicionado e escolha o botão **OK** para fechar a janela de diálogo **Central de Confiabilidade**. 

    ![Diálogo Central de Confiabilidade com o catálogo selecionado](../images/sideload-windows-trust-center-dialog.png)

8. Escolha o botão **OK** para fechar a janela de diálogo **Opções do Word**.

9. Feche e reabra o aplicativo do Office para que as alterações tenham efeito.
    

## <a name="sideload-your-add-in"></a>Fazer o sideload do seu suplemento


1. Coloque o arquivo XML de manifesto de qualquer suplemento que você está testando no catálogo de pasta compartilhada. Observe que você implanta o próprio aplicativo Web em um servidor Web. Certifique-se de especificar a URL no elemento **SourceLocation** do arquivo de manifesto.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. No Excel, Word ou PowerPoint, selecione **Meus Suplementos** na guia **Inserir** da faixa de opções.

3. Escolha **PASTA COMPARTILHADA** na parte superior da caixa de diálogo **Suplementos do Office**.

4. Selecione o nome do suplemento e escolha **OK** para inseri-lo.


## <a name="see-also"></a>Confira também

- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)
    
