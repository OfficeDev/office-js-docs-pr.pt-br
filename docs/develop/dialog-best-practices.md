---
title: Práticas recomendadas e regras para a API da caixa de diálogo do Office
description: Fornece regras e práticas recomendadas para a API de caixa de diálogo do Office, como as práticas recomendadas para um aplicativo de página única (SPA)
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 5e0854137b27d8b8ae33fff8943421cc0c488abe
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292754"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Práticas recomendadas e regras para a API da caixa de diálogo do Office

Este artigo fornece regras, armadilhas e práticas recomendadas para a API de diálogo do Office, incluindo as práticas recomendadas para projetar a interface do usuário de uma caixa de diálogo e usar a API com um aplicativo de página única (SPA)

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com as noções básicas de usar a API de caixa de diálogo do Office, conforme descrito em [usar a API de caixa de diálogo do Office em seus suplementos do Office](dialog-api-in-office-add-ins.md).
> 
> Consulte também [manipulação de erros e eventos com a caixa de diálogo do Office](dialog-handle-errors-events.md).

## <a name="rules-and-gotchas"></a>Regras e dicas

- A caixa de diálogo só pode navegar para URLs HTTPS, não para HTTP.
- A URL passada para o método [displayDialogAsync](/javascript/api/office/office.ui) deve estar no mesmo domínio que o suplemento em si. Ele não pode ser um subdomínio. Mas a página que é passada para ela pode redirecionar para uma página em outro domínio.
- Uma janela hospedeira, que pode ser um painel de tarefas ou o arquivo de [função](../reference/manifest/functionfile.md) sem interface do usuário de um suplemento, pode ter apenas uma caixa de diálogo aberta por vez.
- Apenas duas APIs do Office podem ser chamadas na caixa de diálogo:
  - A função [messageParent](/javascript/api/office/office.ui#messageparent-message-) .
  - `Office.context.requirements.isSetSupported` (Para obter mais informações, consulte [especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md).)
- A função [messageParent](/javascript/api/office/office.ui#messageparent-message-) pode ser chamada apenas de uma página no mesmo domínio que o suplemento.

## <a name="best-practices"></a>Práticas recomendadas

### <a name="avoid-overusing-dialog-boxes"></a>Evitar a superutilização de caixas de diálogo

Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso. Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias. Por exemplo, confira o exemplo [Suplemento do Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

### <a name="designing-a-dialog-box-ui"></a>Criar uma interface do usuário de caixa de diálogo

Para obter as práticas recomendadas no design da caixa de diálogo, consulte [caixas de diálogo em suplementos do Office](../design/dialog-boxes.md).

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a>Tratamento de bloqueadores de pop-up com o Office na Web

A tentativa de exibir uma caixa de diálogo ao usar o Office na Web pode fazer com que o bloqueador de pop-up do navegador bloqueie a caixa de diálogo. O Office na Web tem um recurso que permite que as caixas de diálogo do suplemento sejam uma exceção para o bloqueador de pop-ups do navegador. Quando o código chamar o `displayDialogAsync` método, o Office na Web abrirá um prompt semelhante ao seguinte.

![O prompt que um suplemento pode gerar para evitar bloqueadores de pop-ups no navegador.](../images/dialog-prompt-before-open.png)

Se o usuário escolher **permitir**, a caixa de diálogo do Office será aberta. Se o usuário escolher **ignorar**, o prompt será fechado e a caixa de diálogo do Office não será aberta. Em vez disso, o `displayDialogAsync` método retorna o erro 12009. O código deve capturar esse erro e fornecer uma experiência alternativa que não requer uma caixa de diálogo ou exibir uma mensagem para o usuário que avisa que o suplemento exige que eles permitam a caixa de diálogo. (Para saber mais sobre 12009, confira [erros de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)

Se, por qualquer motivo, você deseja desativar esse recurso, o código deve ser recusado. Ele faz essa solicitação com o objeto [dialogoptions](/javascript/api/office/office.dialogoptions) que é passado para o `displayDialogAsync` método. Especificamente, o objeto deve incluir `promptBeforeOpen: false` . Quando essa opção é definida como false, o Office na Web não solicitará que o usuário permita que o suplemento abra uma caixa de diálogo e a caixa de diálogo do Office não será aberta.

### <a name="do-not-use-the-_host_info-value"></a>Não usar o \_ valor de \_ informações do host

O Office adiciona automaticamente um parâmetro de consulta chamado `_host_info` à URL que é transmitida para `displayDialogAsync` . Ele é acrescentado após os parâmetros de consulta personalizados, se houver. Ele não é acrescentado a quaisquer URLs subsequentes às quais a caixa de diálogo navega. A Microsoft pode alterar o conteúdo desse valor ou removê-lo totalmente, portanto seu código não deve lê-lo. O mesmo valor é adicionado ao armazenamento da sessão da caixa de diálogo. Novamente, *seu código não deve ser lido nem gravado para esse valor*.

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>Práticas recomendadas para usar a API de caixa de diálogo do Office em um SPA

Se seu suplemento usa o roteamento do lado do cliente, como aplicativos de página única (spas usam) normalmente, você tem a opção de passar a URL de uma rota para o método [displayDialogAsync](/javascript/api/office/office.ui) em vez da URL de uma página HTML separada. *Recomendamos que você faça isso pelos motivos apresentados abaixo.*

> [!NOTE]
> Este artigo não é relevante para o roteamento *do lado do servidor* , como em um aplicativo Web baseado em Express.

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>Problemas com o spas usam e a API de caixa de diálogo do Office

A caixa de diálogo do Office está em uma nova janela com sua própria instância do mecanismo JavaScript e, portanto, é o contexto de execução completo. Se você passar uma rota, sua página de base e todos os seus códigos de inicialização e Bootstrap serão executados novamente neste novo contexto, e qualquer variável será definida para seus valores iniciais na caixa de diálogo. Portanto, essa técnica baixa e inicia uma segunda instância do aplicativo na janela da caixa, que anula parcialmente a finalidade de um SPA. Além disso, o código que altera as variáveis na janela da caixa de diálogo não altera a versão do painel de tarefas das mesmas variáveis. Da mesma forma, a janela da caixa de diálogo tem seu próprio armazenamento de sessão, que não é acessível a partir do código no painel de tarefas. A caixa de diálogo e a página host na qual o `displayDialogAsync` foi chamado têm a aparência de dois clientes diferentes para o seu servidor. (Para obter um lembrete sobre o que é uma página de host, consulte [abrir uma caixa de diálogo em uma página de host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).)

Portanto, se você passou uma rota para o `displayDialogAsync` método, você realmente não tem um spa; você teria *duas instâncias do mesmo Spa*. Além disso, grande parte do código na instância do painel de tarefas nunca seria usada nessa instância, e grande parte do código na instância da caixa de diálogo nunca seria usada nessa instância. Seria como ter dois SPAs no mesmo grupo.

#### <a name="microsoft-recommendations"></a>Recomendações da Microsoft

Em vez de passar uma rota do lado do cliente para o `displayDialogAsync` método, recomendamos que você siga um destes procedimentos:

* Se o código que você deseja executar na caixa de diálogo for suficientemente complexo, crie dois spas usam distintos explicitamente; ou seja, têm duas spas usam em pastas diferentes do mesmo domínio. Um SPA é executado na caixa de diálogo e o outro na página host da caixa de diálogo, onde `displayDialogAsync` foi chamado. 
* Na maioria dos cenários, só é necessária uma lógica simples na caixa de diálogo. Nesses casos, o projeto será bastante simplificado, hospedando uma única página HTML, com JavaScript incorporado ou referenciado, no domínio de seu SPA. Passe a URL da página para o método`displayDialogAsync`. Embora isso signifique que você está deviating da ideia literal de um aplicativo de página única; na verdade, você não tem uma única instância de um SPA ao usar a API de diálogo do Office.
