---
title: Empacotar seu suplemento usando o Visual Studio para preparar a publica??o
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e03959294536eeb416a1531d2d281ba83f2d3732
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Empacotar seu suplemento usando o Visual Studio para preparar a publica??o

Seu pacote de Suplemento do Office cont?m um [arquivo de manifesto XML](../develop/add-in-manifests.md) que deve ser usado para publicar o suplemento. Voc? ter? que publicar os arquivos do aplicativo Web do seu projeto separadamente. Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2015

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>Para implantar seu projeto Web usando o Visual Studio 2015

Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2015.

1. No **Gerenciador de Solu??es**, abra o menu de atalho do projeto do suplemento e escolha **Publicar**.
    
    A p?gina **Publicar seu suplemento** ? exibida.
    
2. Na lista suspensa **Perfil atual**, selecione um perfil ou escolha **Novo...** para criar um novo perfil.
    
    > [!NOTE]
    > Um perfil de publica??o especifica o servidor que voc? est? implantando, as credenciais necess?rias para fazer logon no servidor, os bancos de dados para implantar e outras op??es de implanta??o.

    Se voc? escolher **Novo...**, o assistente **Criar perfil de publica??o** ser? exibido. Use esse assistente para importar um perfil de publica??o de um site de hospedagem, como o Microsoft Azure, ou criar um novo perfil e adicionar seu servidor, as credenciais e outras configura??es no procedimento seguinte.
    
    Para mais informa??es sobre como importar perfis de publica??o ou criar novos perfis de publica??o, confira [Criar um Perfil de Publica??o](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. Na p?gina **Publicar seu suplemento**, escolha o link **Implantar seu projeto Web**.
    
    A caixa de di?logo  **Publicar Web** aparece. Para mais informa??es sobre a utiliza??o do desse assistente, veja [Instru??es: Implantar um Projeto da Web usando o On-Click Publishing no Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>Para empacotar seu suplemento usando o Visual Studio 2015

Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2015.

1. Na p?gina **Publicar seu suplemento**, escolha o link **Empacotar o suplemento**.
    
    O assistente **Publicar Suplementos do Office e do SharePoint** ? exibido.
    
2. Na lista suspensa **Onde seu site est? hospedado?**, escolha ou digite a URL do site que hospedar? os arquivos de conte?do do seu suplemento e escolha **Concluir**. 
    
    Voc? deve especificar uma URL que comece com o prefixo HTTPS para concluir este assistente. Se voc? quiser usar um ponto de extremidade HTTP para o site, abra o arquivo de manifesto XML em um editor de texto ap?s criar o pacote e substitua o prefixo HTTPS do site por um prefixo HTTP. 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Os sites do Azure fornecem um ponto de extremidade HTTPS automaticamente.

    O Visual Studio gera os arquivos nos quais voc? precisa publicar seu suplemento e, em seguida, abre a pasta de sa?da de publica??o. 
    
Se voc? planeja enviar seu suplemento ao AppSource, pode escolher o link **Executar uma verifica??o de valida??o** para identificar problemas que possam impedir a aceita??o de seu suplemento. Resolva todos os problemas antes de enviar seu suplemento para a loja.

Agora ? poss?vel carregar seu manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). ? poss?vel encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>Veja tamb?m

- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Disponibilizar suas solu??es no AppSource e no Office](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)
    
