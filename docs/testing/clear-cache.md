---
title: Limpar o cache do Office
description: Saiba como limpar o cache do Office em seu computador.
ms.date: 08/02/2021
localization_priority: Priority
ms.openlocfilehash: 8ae2408b2dbf36a0e5ebbdd863b8ddb49717a144
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773066"
---
# <a name="clear-the-office-cache"></a>Limpar o cache do Office

Você pode remover um suplemento em que foi feito sideload no Windows, Mac ou iOS limpando o cache do Office em seu computador.

Além disso, se você fizer alterações no manifesto do suplemento (por exemplo, atualizar nomes de arquivos de ícones ou texto de comandos de suplemento), deverá limpar o cache do Office e, em seguida, fazer o sideload do suplemento novamente usando o manifesto atualizado. Isso permitirá que o Office renderize o suplemento conforme descrito pelo manifesto atualizado.

> [!NOTE]
> Para remover um suplemento com sideload do Excel, OneNote, PowerPoint ou Word na Web, consulte [Sideload de Suplementos do Office no Office na Web para teste: remover um suplemento de sideload](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in).

## <a name="clear-the-office-cache-on-windows"></a>Limpar o cache do Office no Windows

Para remover todos os suplementos carregados do Excel, Word e PowerPoint, exclua o conteúdo da pasta:

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

Se a pasta a seguir existir, exclua seu conteúdo também.

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

Para remover um suplemento sideload do Outlook, use as etapas descritas em [Suplementos Sideload do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md) para localizar o suplemento na seção **Suplementos Personalizados** da caixa de diálogo que lista seus suplementos instalados. Escolha as reticências (`...`) para o suplemento e, em seguida, escolha **Remover** para remover esse suplemento específico. Se a remoção do suplemento não funcionar, exclua o conteúdo da pasta `Wef` conforme observado anteriormente para Excel, Word e PowerPoint.

Além disso, para limpar o cache do Office no Windows 10 quando o suplemento estiver sendo executado no Microsoft Edge, você pode usar o Microsoft Edge DevTools.

> [!TIP]
> Se você deseja apenas que o suplemento sideloaded reflita as alterações recentes em seus arquivos de origem HTML ou JavaScript, não deve ser necessário limpar o cache. Em vez disso, coloque o foco no painel de tarefas do suplemento (clicando em qualquer lugar no painel de tarefas) e, em seguida, pressione **F5** para recarregar o suplemento.

> [!NOTE]
> Para limpar o cache do Office usando as etapas a seguir, seu suplemento deve ter um painel de tarefas. Se o seu suplemento for um suplemento sem interface de usuário, por exemplo, um que use o recurso [em envio](../outlook/outlook-on-send-addins.md), você precisará adicionar um painel de tarefas ao seu suplemento que use o mesmo domínio para [SourceLocation](../reference/manifest/sourcelocation.md), antes de poder usar as etapas a seguir para limpar o cache.

1. Instalar o [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).

2. Abra seu suplemento no cliente do Office.

3. Execute o Microsoft Edge DevTools.

4. No Microsoft Edge DevTools, abra a guia **Local**. Seu suplemento será listado por nome.

5. Selecione o nome do suplemento para anexar o depurador ao seu suplemento. Uma nova janela do Microsoft Edge DevTools será aberta quando o depurador for anexado ao seu suplemento.

6. Na guia **Network** da nova janela, selecione o botão **Limpar cache**.

    ![Captura de tela do Microsoft Edge DevTools com o botão Limpar cache realçado.](../images/edge-devtools-clear-cache.png)

7. Se concluir essas etapas não produzir o resultado desejado, você também pode selecionar o botão **Sempre atualizar do servidor**.

    ![Captura de tela do Microsoft Edge DevTools com o botão sempre atualizar do servidor realçado.](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>Limpar o cache do Office no Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a>Limpar o cache do Office no iOS

Para limpar o cache do Office no iOS, chame `window.location.reload(true)` a partir do JavaScript no suplemento para forçar um recarregamento. Uma outra alternativa é reinstalar o Office.

## <a name="see-also"></a>Confira também

- [Depurar suplementos do Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md)
- [Realizar sideload de suplementos do Office para teste](sideload-office-add-ins-for-testing.md)
- [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md)
- [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md)
