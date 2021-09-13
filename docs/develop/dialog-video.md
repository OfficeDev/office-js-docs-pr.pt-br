---
title: Use a caixa de diálogo do Office para reproduzir um vídeo
description: Saiba como abrir e reproduzir um vídeo na caixa Office caixa de diálogo
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4765c9d549bfcc52315420342bf99ec88fdebf24
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148997"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>Use a caixa Office caixa de diálogo para mostrar um vídeo

Este artigo explica como reproduzir um vídeo em uma caixa de Office caixa de diálogo do Add-in.

> [!NOTE]
> Este artigo presume que você esteja familiarizado com as noções básicas de uso da caixa de diálogo Office conforme descrito em Usar Office API de diálogo Office em seus Office [Add-ins](dialog-api-in-office-add-ins.md).

Para reproduzir um vídeo em uma caixa de diálogo com a API de Office de diálogo, siga estas etapas.

1. Crie uma página contendo um iframe e nenhum outro conteúdo. A página deve estar no mesmo domínio que a página host. Para um lembrete do que é uma página host, consulte Abrir uma caixa [de diálogo de uma página host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page). No atributo `src` do iframe, aponte para a URL de um vídeo online. O protocolo da URL do vídeo deve ser HTTPS. Neste artigo, chamaremos essa página de "video.dialogbox.html". Veja a seguir um exemplo da marcação.

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. Use uma chamada de `displayDialogAsync` na página host para abrir video.dialogbox.html.
3. Se o suplemento precisar saber quando o usuário fecha a caixa de diálogo, registre um manipulador para o evento `DialogEventReceived` e manipule o evento 12006. Para obter detalhes, consulte [Erros e eventos na caixa Office caixa de diálogo](dialog-handle-errors-events.md).

Para ver um exemplo de um vídeo que está sendo gravado em uma caixa de diálogo, consulte o padrão de [design de placemat de vídeo](../design/first-run-experience-patterns.md#video-placemat).

![Captura de tela mostrando um vídeo que está sendo exibido em uma caixa de diálogo do Excel.](../images/video-placemats-dialog-open.png)
