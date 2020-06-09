---
title: Depurar suplementos do Office em um Mac
description: Saiba como usar um Mac para depurar suplementos do Office
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: bcdf16a858fa613a3248a728b45d68d02cf41a5e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611215"
---
# <a name="debug-office-add-ins-on-a-mac"></a>Depurar suplementos do Office em um Mac

Como os suplementos são desenvolvidos usando HTML e Javascript, são projetados para funcionar em várias plataformas, mas pode haver diferenças sutis em como cada navegador processa o HTML. Este artigo descreve como depurar suplementos em execução em um Mac.

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Depuração com Safari Web Inspetor em um Mac

Se você tiver um suplemento que mostre a interface do usuário em um painel de tarefas ou em um suplemento de conteúdo, o Safari Web Inspector poderá ser usado para depurar um Suplemento do Office.

Para poder depurar Suplementos do Office no Mac, você deverá ter o Mac OS High Sierra E o Mac Office Versão: 16.9.1 (build 18012504) ou posterior. Se você não tiver um build do Office Mac, poderá obter um ingressando no [Programa para Desenvolvedores do Office 365](https://developer.microsoft.com/office/dev-program).

Para iniciar, abra um terminal e defina a propriedade `OfficeWebAddinDeveloperExtras` do aplicativo relevante do Office da seguinte maneira:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

Em seguida, abra o aplicativo do Office e [realize o sideload do seu suplemento](sideload-an-office-add-in-on-ipad-and-mac.md). Clique com o botão direito do mouse no suplemento e você verá a opção **Inspecionar Elemento** no menu de contexto. Marque essa opção e ela exibirá o inspetor, onde você poderá definir os pontos de interrupção e depurar o suplemento.

> [!NOTE]
> Se você estiver tentando usar o inspetor e a caixa de diálogo piscar, atualize o Office para a versão mais recente. Se isso não resolver a oscilação, tente a seguinte solução alternativa:
> 1. Reduza o tamanho da caixa de diálogo.
> 2. Escolha **Inspecionar Elemento**, que será aberto em uma nova janela.
> 3. Redimensione a caixa de diálogo para seu tamanho original.
> 4. Use o inspetor, conforme necessário.

## <a name="clearing-the-office-applications-cache-on-a-mac"></a>Limpar cache do aplicativo do Office em um Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
