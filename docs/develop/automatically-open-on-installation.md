---
title: Abrir automaticamente um painel de tarefas quando um suplemento é instalado
description: Saiba como configurar um Suplemento do Office para abrir automaticamente quando ele estiver instalado.
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d6ff4b8b5b68236d435ec91b2dcbe121f211081d
ms.sourcegitcommit: a32f5613d2bb44a8c812d7d407f106422a530f7a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/14/2022
ms.locfileid: "67674761"
---
# <a name="automatically-open-a-task-pane-when-an-add-in-is-installed"></a>Abrir automaticamente um painel de tarefas quando um suplemento é instalado

Você pode configurar o painel de tarefas do suplemento para ser iniciado imediatamente após a instalação. Esse recurso aumenta o uso. 

Por padrão, os suplementos do painel de tarefas que não  incluem nenhum [](../design/add-in-commands.md) comando de suplemento abrem o painel de tarefas imediatamente após a instalação. No entanto, quando um suplemento tem um ou mais comandos de suplemento, o usuário é notificado sobre o novo suplemento, mas o suplemento não é iniciado automaticamente. Esse comportamento padrão histórico está mudando, portanto, os suplementos que têm comandos de suplemento serão iniciados automaticamente em algumas situações. Além disso, se o suplemento tiver mais de uma página do painel de tarefas, será possível controlar se o suplemento é iniciado na instalação e, nesse caso, qual página será aberta no painel de tarefas.

> [!NOTE]
> 
> - Atualmente, esse recurso está disponível apenas Office na Web. Estamos trabalhando para trazer esse comportamento para outras plataformas, mas atualmente elas ainda exibem o comportamento padrão histórico descrito anteriormente.
> - Esse recurso se aplica somente a suplementos instalados por um usuário final, não a suplementos implantados centralmente.
> - Esse recurso não se aplica a suplementos de conteúdo ou suplementos de Email (Outlook).
> - Esse recurso se aplica somente a suplementos que têm pelo menos um comando de suplemento do tipo ["comando do painel de tarefas"](../design/add-in-commands.md#types-of-add-in-commands).

## <a name="new-behavior"></a>Novo comportamento

O novo comportamento é o seguinte:

- Se o suplemento tiver apenas um comando do painel de [tarefas, a](../design/add-in-commands.md#types-of-add-in-commands) guia da faixa de opções do suplemento será selecionada e o painel de tarefas será aberto automaticamente após a instalação. Você não precisa configurar nada.
- Se o suplemento tiver vários comandos do painel de tarefas e um estiver configurado para ser o padrão (consulte Configurar painel de tarefas [padrão), a](#configure-default-task-pane) guia da faixa de opções do suplemento será selecionada e o painel de tarefas padrão será aberto automaticamente após a instalação.
- Se o suplemento tiver vários comandos do painel de tarefas, mas nenhum estiver configurado para ser o padrão, a guia da faixa de opções do suplemento será selecionada automaticamente após a instalação e um texto explicativo aparecerá perto dele notificando o usuário sobre o novo suplemento, mas nenhum painel de tarefas será aberto. Isso é o mesmo que o comportamento padrão histórico.

> [!NOTE]
> Se, por qualquer motivo, o comando do suplemento que inicia o painel de tarefas não puder ser selecionado manualmente por um usuário na inicialização, como quando ele estiver configurado para [](../design/disable-add-in-commands.md) ser desabilitado na inicialização, ele não será aberto automaticamente, independentemente da configuração. 

## <a name="configure-default-task-pane"></a>Configurar o painel de tarefas padrão

Para designar um painel de tarefas como padrão, adicione um elemento [TaskpaneId](/javascript/api/manifest/action#taskpaneid) **\<Action\>** como o primeiro filho do elemento e defina seu valor como **Office.AutoShowTaskpaneWithDocument**. Apresentamos um exemplo a seguir.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

> [!TIP]
> Se quiser que o suplemento seja iniciado automaticamente sempre que o usuário reabrir o documento, você precisará executar outras etapas de configuração. Para obter detalhes e conselhos sobre quando usar esse recurso, consulte Abrir [automaticamente um painel de tarefas com um documento](automatically-open-a-task-pane-with-a-document.md). 

## <a name="see-also"></a>Confira também

- [Abrir automaticamente um painel de tarefas com um documento](automatically-open-a-task-pane-with-a-document.md)
