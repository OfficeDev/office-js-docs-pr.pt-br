---
title: Depurar suplementos do Office em um Mac
description: Saiba como usar um Mac para depurar Os Complementos do Office.
ms.date: 10/16/2020
localization_priority: Normal
ms.openlocfilehash: b2164e3ed672b2911db6841fad24441b67882204
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237942"
---
# <a name="debug-office-add-ins-on-a-mac"></a>Depurar suplementos do Office em um Mac

Como os suplementos são desenvolvidos usando HTML e Javascript, são projetados para funcionar em várias plataformas, mas pode haver diferenças sutis em como cada navegador processa o HTML. Este artigo descreve como depurar suplementos em execução em um Mac.

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Depuração com Safari Web Inspetor em um Mac

Se você tiver um suplemento que mostre a interface do usuário em um painel de tarefas ou em um suplemento de conteúdo, o Safari Web Inspector poderá ser usado para depurar um Suplemento do Office.

Para poder depurar Os Complementos do Office no Mac, você deve ter o Mac OS High Sierra E o Mac Office versão 16.9.1 (build 18012504) ou posterior. Se você não tiver um build do Office Mac, poderá obter um inando no programa de desenvolvedores do [Microsoft 365.](https://developer.microsoft.com/office/dev-program)

Para iniciar, abra um terminal e defina a propriedade `OfficeWebAddinDeveloperExtras` do aplicativo relevante do Office da seguinte maneira:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > As builds da Mac App Store do Office não são suportadas pelo `OfficeWebAddinDeveloperExtras` sinalizador.

Em seguida, abra o aplicativo do Office e [realize o sideload do seu suplemento](sideload-an-office-add-in-on-ipad-and-mac.md). Clique com o botão direito do mouse no suplemento e você verá a opção **Inspecionar Elemento** no menu de contexto. Marque essa opção e ela exibirá o inspetor, onde você poderá definir os pontos de interrupção e depurar o suplemento.

> [!NOTE]
> Se você estiver tentando usar o inspetor e a caixa de diálogo piscar, atualize o Office para a versão mais recente. Se isso não resolver a oscilação, tente a seguinte solução alternativa:
> 1. Reduza o tamanho da caixa de diálogo.
> 2. Escolha **Inspecionar Elemento**, que será aberto em uma nova janela.
> 3. Redimensione a caixa de diálogo para seu tamanho original.
> 4. Use o inspetor, conforme necessário.

## <a name="clearing-the-office-applications-cache-on-a-mac"></a>Limpar cache do aplicativo do Office em um Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
