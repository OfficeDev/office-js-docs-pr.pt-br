---
title: Fazer sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: e831a1dfbc31ecf06c8b2d78dc1e9a8a4c9dcf01
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/27/2018
ms.locfileid: "21279357"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Fazer sideload de Suplementos do Office para teste usando o **comando sideload**
 >[!NOTE]
>O método "npm run sideload" funciona apenas para suplementos do Excel, Word e PowerPoint.

1. Abra um prompt de comando como administrador.

2. Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.

3. Execute o seguinte comando para iniciar uma instância do servidor da Web local na porta 3000 para servir seu projeto de suplemento: "**npm run start**"

4. Abrir um segundo prompt de comando como administrador.

5. Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.

6. Execute o seguinte comando para inicializar o aplicativo host (por exemplo, Excel, Word) e registre seu suplemento no aplicativo host: "**npm run sideload**"

## <a name="see-also"></a>Confira também

- [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)