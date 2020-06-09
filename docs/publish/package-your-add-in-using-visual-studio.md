---
title: Publicar seu suplemento usando o Visual Studio
description: Como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2019.
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 49b8b53b665b887e4f8dba20e085c3350e7711f8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612049"
---
# <a name="publish-your-add-in-using-visual-studio"></a>Publicar seu suplemento usando o Visual Studio

Seu pacote de Suplemento do Office contém um [arquivo de manifesto XML](../develop/add-in-manifests.md) que deve ser usado para publicar o suplemento. Você terá que publicar os arquivos do aplicativo Web do seu projeto separadamente. Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2019.

> [!NOTE]
> Para saber mais sobre como publicar um Suplemento do Office criado com o gerador Yeoman e desenvolvido com o Código do Visual Studio ou qualquer outro editor, confira [Publicar um suplemento desenvolvido com o Código do Visual Studio](publish-add-in-vs-code.md).

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a>Para implantar seu projeto Web usando o Visual Studio 2019

Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2019.

1. Na guia **Compilar**, escolha **Publicar [Nome do seu suplemento]**.

2. Na janela **Escolha um destino de publicação**, escolha uma das opções de publicação para o seu destino preferido. Cada destino de publicação exige que você inclua mais informações para começar, como um local de pasta ou uma Máquina Virtual do Azure. Depois de especificar um local de publicação e preencher todas as informações necessárias, selecione **Publicar**

    > [!NOTE]
    > A escolha de um destino de publicação especifica o servidor que você está implantando, as credenciais necessárias para fazer logon no servidor, os bancos de dados para implantar e outras opções de implantação.

3. Para obter mais informações sobre as etapas de implantação de cada opção de destino de publicação, confira [Primeiro contato com a implantação no Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a>Para empacotar e publicar seu suplemento usando IIS, FTP ou implantação da Web usando o Visual Studio 2019

Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2019.

1. Na guia **Compilar**, escolha **Publicar [Nome do seu suplemento]**.
2. Na janela **Escolha um destino de publicação**, escolha **IIS, FTP, etc** e selecione **Configurar**. Em seguida, selecione **Publicar**.
3. Será exibido um assistente que o ajudará durante todo o processo. Verifique se o método de publicação é o método preferido, como implantação da Web.
4. Na caixa **URL de destino**, digite a URL do site que hospedará os arquivos de conteúdo do seu suplemento e, em seguida, selecione **Avançar**. Se você pretende enviar seu suplemento ao AppSource, escolha o botão **Validar conexão** para identificar problemas que possam impedir a aceitação do seu suplemento. Você deve resolver todos os problemas antes de enviar seu suplemento para a loja.
5. Confirme as configurações desejadas, incluindo **Opções de publicação de arquivo** e selecione **Salvar**.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.

Agora é possível carregar seu manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). É possível encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a>Confira também

- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-the-office-store)
