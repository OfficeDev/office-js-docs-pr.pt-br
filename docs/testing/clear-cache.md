---
title: Limpar o cache do Office
description: Saiba como limpar o cache do Office em seu computador.
ms.date: 03/11/2022
ms.localizationpriority: high
ms.openlocfilehash: 87cffbe8d28961f8469fbe149ece029bcaaa481d
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484034"
---
# <a name="clear-the-office-cache"></a>Limpar o cache do Office

Para remover um suplemento que você tenha carregado anteriormente de lado no Windows, Mac ou iOS, você precisa limpar o cache do Office em seu computador.

Além disso, se você fizer alterações no manifesto do seu suplemento (por exemplo, atualizar nomes de arquivos de ícones ou texto de comandos do suplemento), você deve limpar o cache do Office e depois recarregar novamente o suplemento usando um manifesto atualizado. Isso permite que o Office apresente o suplemento como é descrito pelo manifesto atualizado.

> [!NOTE]
> Para remover um suplemento com sideload do Excel, OneNote, PowerPoint ou Word na Web, consulte [Sideload de Suplementos do Office no Office na Web para teste: remover um suplemento de sideload](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in).

## <a name="clear-the-office-cache-on-windows"></a>Limpar o cache do Office no Windows

Existem três maneiras de limpar o cache do Office em um computador Windows: automaticamente, manualmente e usando as ferramentas de desenvolvedor do Microsoft Edge. Os métodos são descritos nas subseções a seguir.

### <a name="automatically"></a>Automaticamente

Este método é recomendado para computadores de desenvolvimento de suplementos. Se a versão do Office no Windows for 2108 ou posterior, as etapas a seguir configuram o cache do Office para ser limpo na próxima vez que o Office for reaberto.

> [!NOTE]
> O método automático não é suportado para Outlook.

1. Na faixa de opções de qualquer host do Office, exceto o Outlook, navegue até **Arquivo** > **Opções** > **Central de Confiabilidade** > **Configurações da Central de Confiabilidade** > **Catálogos de Complementos Confiáveis**.
1. Marque a caixa de seleção **Da próxima vez que o Office iniciar, limpe o cache de todos os suplementos da Web iniciados anteriormente**.

### <a name="manually"></a>Manualmente

O método manual para Excel, Word e PowerPoint é diferente do Outlook.

#### <a name="manually-clear-the-cache-in-excel-word-and-powerpoint"></a>Limpar manualmente o cache no Excel, Word e PowerPoint

Para remover todos os suplementos com sideload de Excel, Word e PowerPoint, exclua o conteúdo da pasta a seguir.

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

Se a pasta a seguir existir, exclua seu conteúdo também.

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

#### <a name="manually-clear-the-cache-in-outlook"></a>Limpar manualmente o cache no Outlook

Para remover um suplemento sideload do Outlook, use as etapas descritas em [Suplementos de sideload do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md) para localizar o suplemento na seção **Suplementos personalizados** da caixa de diálogo caixa que lista seus suplementos instalados. Escolha as reticências (`...`) para o suplemento e escolha **Remover** para remover esse suplemento específico. Se a remoção do suplemento não funcionar, exclua o conteúdo da pasta `Wef` conforme observado anteriormente para Excel, Word e PowerPoint.

### <a name="using-the-microsoft-edge-developer-tools"></a>Usando as ferramentas de desenvolvedor do Microsoft Edge

Para limpar o cache do Office no Windows 10 quando o suplemento estiver em execução no Microsoft Edge, você pode usar o Microsoft Edge DevTools.

> [!TIP]
> Se você deseja apenas que o suplemento sideloaded reflita as alterações recentes em seus arquivos de origem HTML ou JavaScript, não deve ser necessário limpar o cache. Em vez disso, coloque o foco no painel de tarefas do suplemento (clicando em qualquer lugar no painel de tarefas) e, em seguida, pressione **Ctrl+F5** para recarregar o suplemento.

> [!NOTE]
> Para limpar o cache do Office usando as etapas a seguir, seu suplemento deve ter um painel de tarefas. Se o seu suplemento for um suplemento sem interface do usuário - por exemplo, um que usa o recurso [ao enviar](../outlook/outlook-on-send-addins.md) - você precisará adicionar um painel de tarefas ao seu suplemento que usa o mesmo domínio para [SourceLocation](/javascript/api/manifest/sourcelocation), antes de usar as etapas a seguir para limpar o cache.

1. Instalar o [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).

2. Abra seu suplemento no cliente do Office.

3. Execute o Microsoft Edge DevTools.

4. No Microsoft Edge DevTools, abra a guia **Local**. Seu suplemento será listado por nome.

5. Selecione o nome do suplemento para anexar o depurador ao seu suplemento. Uma nova janela do Microsoft Edge DevTools será aberta quando o depurador for anexado ao seu suplemento.

6. Na guia **Rede** da nova janela, selecionar **Limpar cache**.

    ![Captura de tela do Microsoft Edge DevTools com o botão Limpar cache realçado.](../images/edge-devtools-clear-cache.png)

7. Se concluir estas etapas não produzir o resultado desejado, tente selecionar **Sempre atualizar a partir do servidor**.

    ![Captura de tela do Microsoft Edge DevTools com o botão sempre atualizar do servidor realçado.](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>Limpar o cache do Office no Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a>Limpar o cache do Office no iOS

Para limpar o cache do Office no iOS, chame `window.location.reload(true)` do JavaScript no suplemento para forçar uma recarga. Como alternativa, reinstale o Office.

## <a name="see-also"></a>Veja também

- [Solucionar erros de desenvolvimento com Suplementos do Office](troubleshoot-development-errors.md)
- [Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Depurar suplementos usando ferramentas de desenvolvedor para Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
- [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
- [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md)
- [Realizar sideload de suplementos do Office para teste](sideload-office-add-ins-for-testing.md)
- [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md)
- [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md)
