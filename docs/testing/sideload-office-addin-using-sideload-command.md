---
title: Realizar sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: dfa231374133ad857554afaf343362f1415788f4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449964"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Realizar sideload de Suplementos do Office usando o **comando sideload**
 >[!NOTE]
>O método "npm executar sideload" só funciona para Word, Excel e PowerPoint suplementos executados no Windows; e somente para projetos que foi criado com a ferramenta [ **yo office** ](https://github.com/OfficeDev/generator-office) e que têm um `sideload` script na `scripts` seção do arquivo package.json. (Projetos que foram criados com versões anteriores do **yo office** também não tem esse script.) Se o projeto foi criado com o Visual Studio ou não tem o script sideload, você pode sideload no Windows com o método descrito [Sideload um suplemento do Office em um compartilhamento de rede](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
>
> Se não estiver testando um suplemento do Word, do Excel ou do PowerPoint no Windows, confira um dos tópicos a seguir para fazer sideload do suplemento:
> 
> - [Sideload de suplementos do Office para teste no Office Online](sideload-office-add-ins-for-testing.md)
> - [Sideload suplementos do Office para teste em um iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [Realizar sideload de suplementos do Outlook para teste](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. Abra um prompt de comando como administrador.

2. Altere os diretórios na raiz da pasta em um projeto.

3. Execute o seguinte comando para iniciar uma instância do servidor local da web na porta 3000 para atender a seu projeto do suplemento: "**npm executar início**"

4. Abra um segundo prompt de comando como administrador.

5. Altere os diretórios na raiz da pasta em um projeto.

6. Execute o seguinte comando para inicializar o aplicativo de host (por exemplo, o Excel, Word) e inscreva-se o suplemento no aplicativo do host: "**npm executar sideload**"

## <a name="see-also"></a>Confira também

- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)
