---
title: Realizar sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 38aa74963ca750d65e4be7bb17745a59eeed0c83
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126887"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Realizar sideload de Suplementos do Office usando o comando sideload
 
> [!NOTE]
> A técnica de sideload descrita neste artigo é válida somente para:
> 
> - Suplementos do Excel, Word e PowerPoint executados no Windows
> 
> - Os projetos de suplemento que foram criados com o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) e que possuem um script `sideload` na seção `scripts` do arquivo package.json. (Projetos que foram criados com as versões anteriores do gerador Yeoman para Suplementos do Office não possuirão este script.)
 
Para realizar o sideload no seu suplemento usando o script `sideload` que o gerador Yeoman para Suplementos do Office fornece, conclua as seguintes etapas:

1. Abra um prompt de comando como administrador.

2. Altere os diretórios na raiz da pasta em um projeto.

3. Execute o seguinte comando para iniciar uma instância do servidor local da web na porta 3000 para atender ao seu projeto de suplemento: `npm run start`

4. Abra um segundo prompt de comando como administrador.

5. Altere os diretórios na raiz da pasta em um projeto.

6. Execute o seguinte comando para inicializar o aplicativo de host (por exemplo, o Excel ou o Word) e registrar o seu suplemento no aplicativo do host: `npm run sideload`

Se o seu projeto de suplemento foi criado com o Visual Studio ou não possui o script sideload, você pode realizar o sideload no Windows usando o método descrito em [Realizar Sideload em um Suplemento do Office a partir de um compartilhamento de rede](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

Se você não estiver testando um suplemento do Word, do Excel ou do PowerPoint no Windows, confira um dos tópicos a seguir para obter informações sobre como realizar o sideload no seu suplemento:
 
- [Realizar sideload de suplementos do Office no Office Online para teste](sideload-office-add-ins-for-testing.md)
- [Sideload suplementos do Office para teste em um iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Realizar sideload de suplementos do Outlook para teste](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a>Confira também

- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)
