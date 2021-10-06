---
title: Testar e depurar suplementos do Office
description: Aprenda a testar e depurar Suplementos do Office.
ms.date: 09/24/2021
ms.localizationpriority: high
ms.openlocfilehash: 1a47c9e5b4d45dc2a20f7fd5dd04a384eea5cba3
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138594"
---
# <a name="test-and-debug-office-add-ins"></a>Testar e depurar suplementos do Office

Esta seção contém orientações sobre testes, depuração de bugs e solução de problemas em Suplementos do Office.

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>Testar plataforma cruzada e para várias versões do Office

Os Suplementos do Office são executados em grandes plataformas, então é necessário testar um suplemento em todas as plataformas em que seus usuários podem estar executando o Office. Isso normalmente inclui o Office na Web, Office no Windows (tanto assinatura como compra avulsa), Office no Mac, Office no iOS e (para suplementos do Outlook) Office no Android. No entanto, pode haver algumas situações em que você tem certeza de que nenhum de seus usuários estará trabalhando em algumas plataformas. Por exemplo, se você estiver criando um suplemento para uma empresa que exige que seus usuários trabalhem com computadores Windows e assinatura do Office, não será necessário testar o Office no Mac ou o Windows de compra avulsa.

> [!NOTE]
> Em computadores Windows, a versão do Windows e do Office determinarão qual controle de navegador será usado pelos suplementos. Para obter mais informações, veja [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Os suplementos comercializados pelo AppSource passam por um processo de validação que inclui testes em todas as plataformas. Além disso, os suplementos são testados para o Office na Web em todos os principais navegadores modernos, incluindo o Microsoft Edge (WebView2 baseado em Chromium), Chrome e Safari. Teste adequadamente nessas plataformas e navegadores antes de enviar ao AppSource. Para obter mais informações sobre validação, veja [Políticas de certificação de marketplace comercial](/legal/marketplace/certification-policies), principalmente a [seção 1120.3](/legal/marketplace/certification-policies#11203-functionality) e a [página de aplicativo e disponibilidade do Suplemento do Office](../overview/office-add-in-availability.md).
>
> O AppSource não usa o Internet Explorer ou a versão herdada do Microsoft Edge (WebView1) para testar suplementos no Office na Web. Mas se um número significativo de seus usuários usará o Edge herdado para abrir o Office na Web, você deve testar com ele. (O Office na Web não abre no Internet Explorer, portanto você não pode e não precisa testar o Office na Web com o Internet Explorer.) Para obter mais informações, consulte [Suporte ao Internet Explorer 11](../develop/support-ie-11.md) e [Solução de problemas do Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues). O Office ainda oferece suporte a esses navegadores para runtimes de suplementos, portanto, se você acha que encontrou um bug na forma como os suplementos são executados neles, crie um problema para o repositório [office js](https://github.com/OfficeDev/office-js/issues/new/choose).

## <a name="sideload-an-office-add-in-for-testing"></a>Fazer sideload de suplemento para teste

Você pode usar o sideload para instalar um Suplemento do Office para teste sem precisar primeiro colocá-lo em um catálogo de suplementos. O procedimento para sideload de um suplemento varia de acordo com a plataforma e, em alguns casos, por produto também. Os artigos a seguir descrevem como realizar sideload de Suplementos do Office em uma plataforma específica ou em um produto específico.

- [Fazer sideload de Suplementos do Office no Windows](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Realizar sideload de suplementos do Office no Office na Web](sideload-office-add-ins-for-testing.md)

- [Fazer sideload de Suplementos do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md)

- [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a>Depurar um suplemento do Office

O procedimento para depurar um Suplemento do Office também varia de acordo com a plataforma. Cada um dos artigos a seguir descreve como depurar suplementos do Office em uma plataforma específica.

- [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)

- [Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10](debug-add-ins-using-f12-developer-tools-on-windows.md)

- [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md)

- [Depurar Suplementos do Office em um Mac](debug-office-add-ins-on-ipad-and-mac.md)

- [Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a>Validar o manifesto de suplemento do Office

Confira as informações sobre como validar o arquivo de manifesto que descreve os suplementos do Office e solucionar problemas com o arquivo de manifesto em [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md).

## <a name="troubleshoot-user-errors"></a>Solucionar problemas de erros de usuário

Confira informações sobre como solucionar problemas comuns que os usuários podem encontrar em seu suplemento do Office em [Solucionar erros de usuários com os suplementos do Office](testing-and-troubleshooting.md)
