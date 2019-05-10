---
title: Visualizadores Web usados por Suplementos do Office
description: ''
ms.date: 05/03/2019
localization_priority: Priority
ms.openlocfilehash: 632f62cbc02917d9e28ab260f3710498156194db
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33630402"
---
# <a name="web-viewers-used-by-office-add-ins"></a>Visualizadores Web usados por Suplementos do Office

Visto que os Suplementos do Office são aplicativos da Web, eles precisam de um visualizador de páginas da Web para exibir as páginas de HTML do aplicativo da Web e um mecanismo JavaScript para executar o JavaScript. Ambos são fornecidos por um navegador instalado no computador do usuário.

Qual navegador é usado depende do:

- Sistema operacional do computador.
- Se o suplemento está em execução no Office Online, no Office 365 ou no Office 2013 sem assinatura ou posterior.

A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.

|**SO / Plataforma**|**Navegador**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office Online|O navegador no qual o Office Online está aberto.|
|Mac|Safari|
|iOS|Safari|
|Android|Chrome|
|Windows / Office 2013 sem assinatura ou posterior.|Internet Explorer 11|
|Versão do Windows 10 < 1903 / Office 365|Internet Explorer 11|
|Versão do Windows 10 >= 1903 / versão do Office 365 < 16.0.11629|Internet Explorer 11|
|Versão do Windows 10 >= 1903 / versão do Office 365 >= 16.0.11629|Edge\*|

\*Quando o Edge está sendo usado, o Narrador do Windows 10 (também conhecido como «leitor de tela») lê a marca `<title>` na página que abre no painel de tarefas. Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.

> [!IMPORTANT]
> O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5. Se qualquer um dos usuários de suplemento tiverem plataformas com Internet Explorer 11, para que seja possível usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você precisará fazer o transpile do seu JavaScript para o ES5 ou usar um polyfill. Além disso, o Internet Explorer 11 não oferece suporte para alguns recursos de HTML 5, como mídia, gravação e localização.

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

## <a name="see-also"></a>Confira também

- [Requisitos para a Execução de Suplementos do Office](requirements-for-running-office-add-ins.md)
