---
title: Empacotar seu suplemento usando o Visual Studio para preparar a publicação | Microsoft Docs
description: Como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2017.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 9233ebed217c9e4cc5def0dace67043f29462296
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451084"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Empacotar seu suplemento usando o Visual Studio para preparar a publicação

Seu pacote de Suplemento do Office contém um [arquivo de manifesto XML](../develop/add-in-manifests.md) que deve ser usado para publicar o suplemento. Você terá que publicar os arquivos do aplicativo Web do seu projeto separadamente. Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2017.

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a>Implantar seu projeto Web usando o Visual Studio 2017

Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2017.

1. No **Gerenciador de Soluções**, abra o menu de atalho do projeto do suplemento e escolha  **Publicar**.

    A página **Publicar seu suplemento** é exibida.

2. Na lista suspensa **Perfil atual**, selecione um perfil ou escolha **Novo...** para criar um novo perfil.

    > [!NOTE]
    > Um perfil de publicação especifica o servidor que você está implantando, as credenciais necessárias para fazer logon no servidor, os bancos de dados para implantar e outras opções de implantação.

    Se você escolher **Novo ...**, o assistente é exibido com a página **Criar perfil de Publicação**. Use esse assistente para importar um perfil de publicação de um site de hospedagem, como o Microsoft Azure, ou criar um novo perfil e adicionar seu servidor, as credenciais e outras configurações no procedimento seguinte.

    Para mais informações sobre como importar perfis de publicação ou criar novos perfis de publicação, veja [Criar um Perfil de Publicação](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).

3. Na página **Publicar seu suplemento**, escolha o link **Implantar seu projeto Web**.

    A caixa de diálogo **Publicar** é exibida. Para mais informações sobre como usar o assistente, veja [Como: implantar um Projeto Web usando a Publicação On-Click no Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).

## <a name="to-package-your-add-in-using-visual-studio-2017"></a>Empacotar seu suplemento usando o Visual Studio 2017

Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2017.

1. Na página **Publicar seu suplemento**, escolha o botão **Empacotar o suplemento**.

    Um assistente é exibido com a página **Empacotar o suplemento**.

2. Na lista suspensa **Onde seu site está hospedado?**, escolha ou digite a URL do site que hospedará os arquivos de conteúdo do seu suplemento e escolha **Concluir**.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.

    O Visual Studio gera os arquivos nos quais você precisa publicar seu suplemento e, em seguida, abre a pasta de saída de publicação.

Se você pretende enviar seu suplemento ao AppSource, escolha o botão **Executar uma verificação de validação** para identificar problemas que possam impedir a aceitação do seu suplemento. Você deve resolver todos os problemas antes de enviar seu suplemento para a loja.

Agora é possível carregar seu manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). É possível encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a>Confira também

- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-the-office-store)
