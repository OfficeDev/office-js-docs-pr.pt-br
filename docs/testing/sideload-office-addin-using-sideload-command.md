---
title: Fazer sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 1ab0277493f2899adb479c2f24b1635a881af3cc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944038"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Fazer sideload de Suplementos do Office para teste usando o **comando sideload**
 >[!NOTE]
>O método "npm run sideload" só funciona para suplementos de Excel, Word e PowerPoint executados no Windows; e somente para projetos de suplemento criados com a [**ferramenta yo office**](https://github.com/OfficeDev/generator-office) e que possuem um script `sideload` na seção de `scripts` do arquivo package.json. (Projetos que foram criados com versões mais antigas do **yo office** também não têm esse script.) Se o seu projeto foi criado com o Visual Studio ou não tem o script sideload, você poderá fazer o sideload dele no Windows com o método descrito em [Fazer sideload de um Suplemento do Office a partir de um compartilhamento de rede](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
>
> Se não estiver testando um suplemento do Word, do Excel ou do PowerPoint no Windows, confira um dos tópicos a seguir para fazer sideload do suplemento:
> 
> - [Sideload de suplementos do Office para teste no Office Online](sideload-office-add-ins-for-testing.md)
> - [Sideload de suplementos do Office para teste em um iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [Sideload de suplementos do Outlook para teste](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. Abra um prompt de comando como administrador.

2. Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.

3. Execute o seguinte comando para iniciar uma instância do servidor da Web local na porta 3000 para servir seu projeto de suplemento: "**npm run start**"

4. Abra um segundo prompt de comando como administrador.

5. Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.

6. Execute o seguinte comando para inicializar o aplicativo host (por exemplo, Excel, Word) e registre seu suplemento no aplicativo host: "**npm run sideload**"

## <a name="see-also"></a>Confira também

- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)