---
title: Fazer sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 90084fad0e79ab8acdf59eaa305825737401c0c8
ms.sourcegitcommit: e094aaa06d9aff3d13f8ffd3429d4a31f0b65b81
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/03/2018
ms.locfileid: "21782823"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Fazer sideload de Suplementos do Office para teste usando o **comando sideload**
 >[!NOTE]
>O método "npm run sideload" funciona apenas para suplementos do Excel, Word e PowerPoint executados no Windows e para projetos de suplementos criados com a ferramenta [**yo office** e](https://github.com/OfficeDev/generator-office) que têm um script `sideload` na seção `scripts` do arquivo package.json. (Projetos criados com versões mais antigas do **yo office** também não têm esse script.) Se o seu projeto foi criado com o Visual Studio ou não tem o script de sideload, você pode fazer o sideload dele no Windows com o método descrito em [Fazer o sideload de um suplemento do Office a partir de um compartilhamento de rede](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
>
> Se não estiver testando um suplemento do Word, do Excel ou do PowerPoint no Windows, confira um dos tópicos a seguir para fazer sideload do suplemento:
> 
> - [Sideload de suplementos do Office para teste no Office Online](sideload-office-add-ins-for-testing.md)
> - [Sideload dos suplementos do Office para teste em um iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [Fazer sideload de suplementos do Outlook para teste](../../../../outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. Abra um prompt de comando como administrador.

2. Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.

3. Execute o seguinte comando para iniciar uma instância do servidor da Web local na porta 3000 para servir seu projeto de suplemento: "**npm run start**"

4. Abra um segundo prompt de comando como administrador.

5. Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.

6. Execute o seguinte comando para inicializar o aplicativo host (por exemplo, Excel, Word) e registre seu suplemento no aplicativo host: "**npm run sideload**"

## <a name="see-also"></a>Veja também

- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)