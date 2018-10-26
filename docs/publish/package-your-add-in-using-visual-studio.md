---
title: Empacote seu suplemento usando o Visual Studio para preparar a publicação | Microsoft Docs
description: Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2017.
ms.date: 01/25/2018
ms.openlocfilehash: 3515f88e41bc5f0af62a3b043beae5177f3291ac
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681760"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Empacote seu suplemento usando o Visual Studio para preparar a publicação

O pacote do Suplemento do Office contém um [arquivo de manifesto](../develop/add-in-manifests.md) XML que você utilizará para publicar o suplemento. Você terá que publicar os arquivos do aplicativo da Web do seu projeto separadamente. Este artigo descreve como implantar o seu projeto Web e empacotar seu suplemento usando o Visual Studio 2017.

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a>Para implantar seu projeto Web usando o Visual Studio 2017

Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2017.

1. No **Gerenciador de Soluções**, abra o menu de atalho do projeto do suplemento e escolha **Publicar**.
    
    A página **Publicar seu suplemento** é exibida.
    
2. Na lista suspensa **Perfil atual**, selecione um perfil ou escolha **Novo...** para criar um novo perfil.
    
    > [!NOTE]
    > Um perfil de publicação especifica o servidor de implantação, as credenciais necessárias para fazer logon no servidor, os bancos de dados a serem implantados e outras opções de implantação.

    Se você escolher **Novo...**, o assistente será exibido com a página **Criar perfil de publicação**. Você pode usar esse assistente para importar um perfil de publicação de um provedor de hospedagem de sites da Web, como o Microsoft Azure, ou criar um novo perfil e adicionar seu servidor, as credenciais e outras configurações no próximo procedimento.
    
    Para obter mais informações sobre como importar perfis de publicação ou criar novos perfis de publicação, confira [Criar um Perfil de Publicação](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).
    
3. Na página **Publicar seu suplemento**, escolha o link **Implantar seu projeto Web**.
    
    A caixa de diálogo  **Publicar** será exibida. Para obter mais informações sobre como usar esse assistente, consulte [Tutorial: Implantar um projeto Web usando o On-Click Publishing no Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a>Para empacotar seu suplemento usando o Visual Studio 2017

Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2017.

1. Na página **Publicar seu suplemento**, escolha o botão**Empacotar o suplemento**.
    
    Será exibido o assistente com a página **Empacotar o suplemento**.
    
2. Na caixa  **Onde seu site da Web está hospedado?**, insira a URL do site da Web que hospedará os arquivos de conteúdo do seu suplemento e escolha **Concluir**.
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Os sites da Web do Azure fornecem automaticamente um ponto de extremidade HTTPS.

    O Visual Studio gera os arquivos que você precisa para publicar seu suplemento e, em seguida, abre a pasta de saída da publicação.
    
Se você planeja enviar o suplemento para o AppSource, pode escolher o botão **Executar uma verificação de validação** para identificar problemas que possam impedir a aceitação do suplemento. Você deve resolver todos os problemas antes de enviar o suplemento para o repositório.

Agora é possível carregar o manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). É possível encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>Confira também

- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Disponibilizar suas soluções no AppSource e no Office](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
