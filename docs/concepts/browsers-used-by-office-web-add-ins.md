---
title: Navegadores usados pelos Suplementos do Office
description: Especifica como o sistema operacional e a versão do Office determinam o navegador que é usado pelos suplementos do Office.
ms.date: 09/25/2019
localization_priority: Priority
ms.openlocfilehash: b5d7198e556f020bccdf7ba1e0a0fcffa3a9171b
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235292"
---
# <a name="browsers-used-by-office-add-ins"></a>Navegadores usados pelos Suplementos do Office

Os suplementos do Office são aplicativos Web exibidos usando iFrames durante a execução do Office na Web e no uso de controles de navegador incorporados no Office para clientes desktops e móveis. Os suplementos também precisam de um mecanismo JavaScript para executar o JavaScript. O navegador incorporado e o mecanismo são fornecidos por um navegador instalado no computador do usuário.

Qual navegador é usado depende do:

- Sistema operacional do computador.
- Se o suplemento está em execução no Office na Web, no Office 365 ou no Office 2013 sem assinatura ou posterior.

A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.

|**SO / Plataforma**|**Navegador**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office na Web|O navegador no qual o Office está aberto.|
|Mac|Safari|
|iOS|Safari|
|Android|Chrome|
|Windows / Office 2013 sem assinatura ou posterior.|Internet Explorer 11|
|Versão do Windows 10 < 1903 / Office 365|Internet Explorer 11|
|Versão do Windows 10 >= 1903 / versão do Office 365 < 16.0.11629|Internet Explorer 11|
|Versão do Windows 10 >= 1903 / versão do Office 365 >= 16.0.11629|Microsoft Edge\*|

\*Quando o Microsoft Edge está sendo usado, o Windows 10 Narrator (às vezes chamado de "leitor de tela") lê a marcação `<title>` na página que é aberta no painel de tarefas. Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.

> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se qualquer um dos usuários de suplemento tiverem plataformas com Internet Explorer 11, para que seja possível usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você precisará fazer o transpile do seu JavaScript para o ES5 ou usar um polyfill. Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.

> [!NOTE]
> Até que eles estejam disponíveis, você precisará ser um Windows Insider para obter a versão 1903 do Windows ou superior, e ser um Office Insider para obter a versão 16.0.11629 do Office ou superior.
>
> Para participar do programa Windows Insider:
> 
> 1. Vá até [Windows Insider](https://insider.windows.com) e clique no link para participar do Windows Insider.
> 2. Você será direcionado para uma página com instruções sobre como usar as Configurações do Windows para habilitar as compilações de visualização do Windows. Siga as instruções. Quando for selecionar a velocidade das atualizações, escolha a opção mais rápida.
>
> Para participar do programa Office Insider:
> 
> 1. Vá até [Introdução ao Programa Office Insider](https://insider.office.com/join).
> 2. Siga as instruções na página para participar. Quando solicitado a especificar um canal, selecione Insider.

## <a name="troubleshooting-microsoft-edge-issues"></a>Solucionar problemas do Microsoft Edge

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>Barra de rolagem não aparece no painel de tarefas

Por padrão, as barras de rolagem no Microsoft Edge estão ocultas até que você tenha passado. Para garantir que a barra de rolagem fique sempre visível, o estilo de CSS que se aplica ao elemento `<body>` das páginas no painel de tarefas deve incluir a propriedade [(-ms- reoverflow-style)](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) e deve ser definida como `scrollbar`. 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Ao depurar com o Microsoft Edge DevTools, o suplemento falha ou recarrega

A definição de pontos de interrupção nas [DevTools do Microsoft Edge](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) pode fazer o Office pensar que o suplemento está travado. Ele recarrega automaticamente o suplemento quando isso acontece. Para evitar isso, adicione a seguinte chave do registro e valor ao computador de desenvolvimento `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`:.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Quando o suplemento tentar abrir, o erro “ADD-IN ERROR não é possível abrir este suplemento a partir do localhost" acontece

Uma causa conhecida é que o Microsoft Edge exige que o localhost tenha uma isenção de auto-retorno no computador de desenvolvimento. Siga as instruções em [não é possível abrir o suplemento do localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).


## <a name="see-also"></a>Confira também

- [Requisitos para a Execução de Suplementos do Office](requirements-for-running-office-add-ins.md)
