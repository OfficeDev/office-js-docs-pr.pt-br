---
title: Práticas recomendadas e regras para a API da caixa de diálogo do Office
description: Fornece regras e práticas recomendadas para a API de caixa de diálogo do Office, como práticas recomendadas para um aplicativo de página única (SPA)
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 4359d116e9720255278c5b3f543b135013c7e76c
ms.sourcegitcommit: 7cd501d0fdbbd4636bd08647b638dd5ca4c7c630
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/17/2021
ms.locfileid: "50282979"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Práticas recomendadas e regras para a API da caixa de diálogo do Office

Este artigo fornece regras, conversas e práticas recomendadas para a API de caixa de diálogo do Office, incluindo práticas recomendadas para projetar a interface do usuário de uma caixa de diálogo e usar a API com um aplicativo de página única (SPA)

> [!NOTE]
> Este artigo pressupõe que você está familiarizado com as noções básicas de uso da API de caixa de diálogo do Office, conforme descrito em Usar a API de caixa de diálogo do Office em seus [Complementos do Office.](dialog-api-in-office-add-ins.md)
> 
> Consulte também [Manipulando erros e eventos com a caixa de diálogo do Office.](dialog-handle-errors-events.md)

## <a name="rules-and-gotchas"></a>Regras e dicas

- A caixa de diálogo só pode navegar até URLs HTTPS, não HTTP.
- A URL passada para [o método displayDialogAsync](/javascript/api/office/office.ui) deve estar exatamente no mesmo domínio que o próprio complemento. Não pode ser um subdomínio. Mas a página que é passada para ela pode redirecionar para uma página em outro domínio.
- Uma janela de host, que pode ser um [](../reference/manifest/functionfile.md) painel de tarefas ou o arquivo de função sem interface do usuário de um comando de complemento, pode ter apenas uma caixa de diálogo aberta por vez.
- Somente duas APIs do Office podem ser chamadas na caixa de diálogo:
  - A [função messageParent.](/javascript/api/office/office.ui#messageparent-message-)
  - `Office.context.requirements.isSetSupported`(Para saber mais, confira Especificar [aplicativos do Office e requisitos de API.)](specify-office-hosts-and-api-requirements.md)
- A [função messageParent](/javascript/api/office/office.ui#messageparent-message-) só pode ser chamada de uma página no mesmo domínio que o próprio complemento.

## <a name="best-practices"></a>Práticas recomendadas

### <a name="avoid-overusing-dialog-boxes"></a>Evitar o excesso de caixas de diálogo

Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso. Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias. Para ver um exemplo de um painel de tarefas com guias, confira o exemplo [do JavaScript SalesTracker para o](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) Complemento do Excel.

### <a name="designing-a-dialog-box-ui"></a>Projetando uma interface do usuário de caixa de diálogo

Para saber as práticas recomendadas no design da caixa de diálogo, confira [Caixas de diálogo nos Complementos do Office.](../design/dialog-boxes.md)

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a>Tratamento de bloqueadores de pop-up com o Office na Web

Tentar exibir uma caixa de diálogo ao usar o Office na Web pode fazer com que o bloqueador de pop-ups do navegador bloqueie a caixa de diálogo. O Office na Web tem um recurso que permite que as caixas de diálogo do seu complemento sejam uma exceção ao bloqueador de pop-ups do navegador. Quando seu código chamar o `displayDialogAsync` método, o Office na Web abrirá um prompt semelhante ao seguinte.

![Screenshot showing the prompt with a brief description and Allow and Ignore buttons that an add-in can generate to avoid in-browser pop-up blockers](../images/dialog-prompt-before-open.png)

Se o usuário escolher Permitir, **a** caixa de diálogo do Office será aberta. Se o usuário escolher **Ignorar**, o prompt será fechado e a caixa de diálogo do Office não será aberta. Em vez disso, `displayDialogAsync` o método retorna o erro 12009. Seu código deve capturar esse erro e fornecer uma experiência alternativa que não exija uma caixa de diálogo ou exibir uma mensagem para o usuário avisando que o complemento exige que ele permita a caixa de diálogo. (Para obter mais sobre 12009, consulte [Erros de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)

Se, por qualquer motivo, você quiser desativar esse recurso, seu código deverá re optá-lo. Ela faz essa solicitação com o [objeto DialogOptions](/javascript/api/office/office.dialogoptions) que é passado para o `displayDialogAsync` método. Especificamente, o objeto deve incluir `promptBeforeOpen: false` . Quando essa opção for definida como falsa, o Office na Web não solicitará que o usuário permita que o complemento abra uma caixa de diálogo e a caixa de diálogo do Office não será aberta.

### <a name="do-not-use-the-_host_info-value"></a>Não use o valor \_ de \_ informações do host

O Office adiciona automaticamente um parâmetro de `_host_info` consulta chamado para a URL que é passada para `displayDialogAsync` . Ele é anexado após os parâmetros de consulta personalizados, se algum. Ele não é anexado a quaisquer URLs subsequentes para as que a caixa de diálogo navega. A Microsoft pode alterar o conteúdo desse valor ou removê-lo completamente, para que seu código não o leia. O mesmo valor é adicionado ao armazenamento da sessão da caixa de diálogo (ou seja, a [propriedade Window.sessionStorage).](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) Novamente, *seu código não deve ler nem gravar nesse valor.*

### <a name="opening-another-dialog-immediately-after-closing-one"></a>Abrir outra caixa de diálogo imediatamente após fechar uma

Você não pode ter mais de uma caixa de diálogo aberta em uma determinada página host, portanto, seu código deve chamar [Dialog.close](/javascript/api/office/office.dialog#close__) em uma caixa de diálogo aberta antes de chamar para `displayDialogAsync` abrir outra caixa de diálogo. O `close` método é assíncrono. Por esse motivo, se você chamar imediatamente após uma chamada de , a primeira caixa de diálogo poderá não ter sido totalmente fechada quando o Office tentar `displayDialogAsync` `close` abrir a segunda. Se isso acontecer, o Office retornará um erro [12007:](dialog-handle-errors-events.md#12007) "A operação falhou porque este complemento já tem uma caixa de diálogo ativa".

O método não aceita um parâmetro de retorno de chamada e não retorna um objeto Promise para que ele não possa ser aguardado com a palavra-chave ou `close` `await` com um `then` método. Por esse motivo, sugerimos a técnica a seguir quando você precisar abrir uma nova caixa de diálogo imediatamente após fechar uma caixa de diálogo: encapsular o código para abrir a nova caixa de diálogo em um método e projetar o método para chamar a si mesmo recursivamente se a chamada de retorno `displayDialogAsync` `12007` . Apresentamos um exemplo a seguir.

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        openSecondDialog();
      }
      else {
         // Handle errors
      }
    }
  );
}
 
function openSecondDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
    (result) => {
      if(result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          openSecondDialog(); // Recursive call
        }
        else {
         // Handle other errors
        }
      }
    }
  );
}
```

Como alternativa, você pode forçar o código a pausar antes que ele tente abrir a segunda caixa de diálogo usando o [método setTimeout.](https://www.w3schools.com/jsref/met_win_settimeout.asp) Apresentamos um exemplo a seguir.

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        setTimeout(() => { 
          Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
             (result) => { /* callback body */ }
          );
        }, 1000);
      }
      else {
         // Handle errors
      }
    }
  );
}
```

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>Práticas recomendadas para usar a API de caixa de diálogo do Office em um SPA

Se o seu complemento usa o roteamento do lado do cliente, como os aplicativos de página única (SPAs) normalmente fazem, você tem a opção de passar a URL de uma rota para o método [displayDialogAsync](/javascript/api/office/office.ui) em vez da URL de uma página HTML separada. *Não recomendamos isso pelos motivos abaixo.*

> [!NOTE]
> Este artigo não é relevante para *o roteamento* do lado do servidor, como em um aplicativo Web baseado em Express.

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>Problemas com SPAs e a API de caixa de diálogo do Office

A caixa de diálogo do Office está em uma nova janela com sua própria instância do mecanismo JavaScript e, portanto, é seu próprio contexto de execução completo. Se você passar uma rota, sua página base e todo o código de inicialização e inicialização serão executados novamente nesse novo contexto, e todas as variáveis serão definidas com seus valores iniciais na caixa de diálogo. Portanto, essa técnica baixa e inicia uma segunda instância do seu aplicativo na janela da caixa, o que anuncia parcialmente a finalidade de um SPA. Além disso, o código que altera variáveis na janela da caixa de diálogo não altera a versão do painel de tarefas das mesmas variáveis. Da mesma forma, a janela da caixa de diálogo tem seu próprio armazenamento de sessão (a [propriedade Window.sessionStorage),](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) que não é acessível a partir do código no painel de tarefas. A caixa de diálogo e a página host na qual `displayDialogAsync` foi chamada parecem dois clientes diferentes para seu servidor. (Para um lembrete do que é uma página host, consulte [Abrir uma caixa de diálogo em uma página host.)](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)

Portanto, se você passar uma rota para o método, não terá realmente um SPA; você teria duas `displayDialogAsync` *instâncias do mesmo SPA.* Além disso, grande parte do código na instância do painel de tarefas nunca seria usada nessa instância e grande parte do código na instância da caixa de diálogo nunca seria usada nessa instância. Seria como ter dois SPAs no mesmo grupo.

#### <a name="microsoft-recommendations"></a>Recomendações da Microsoft

Em vez de passar uma rota do lado do cliente para o método, recomendamos `displayDialogAsync` que você faça um dos seguintes:

* Se o código que você deseja executar na caixa de diálogo for suficientemente complexo, crie dois SPAs diferentes explicitamente; ou seja, ter dois SPAs em pastas diferentes do mesmo domínio. Um SPA é executado na caixa de diálogo e o outro na página host da caixa de diálogo onde `displayDialogAsync` foi chamado. 
* Na maioria dos cenários, apenas a lógica simples é necessária na caixa de diálogo. Nesses casos, seu projeto será bastante simplificado hospedando uma única página HTML, com JavaScript incorporado ou referenciado, no domínio de seu SPA. Passe a URL da página para o método`displayDialogAsync`. Isso significa que você está desviando da ideia literal de um aplicativo de página única; na verdade, você não tem uma única instância de um SPA quando está usando a API de caixa de diálogo do Office.
