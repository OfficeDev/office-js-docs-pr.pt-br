---
title: Práticas recomendadas e regras para a API da caixa de diálogo do Office
description: Fornece regras e práticas recomendadas para a API Office caixa de diálogo, como práticas recomendadas para um SPA (aplicativo de página única).
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: c4594bc8636bd40b4b2511e3faa4fd879c5b2f10
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628051"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Práticas recomendadas e regras para a API da caixa de diálogo do Office

Este artigo fornece regras, gotchas e práticas recomendadas para a API de diálogo Office, incluindo práticas recomendadas para projetar a interface do usuário de uma caixa de diálogo e usar a API em um SPA (aplicativo de página única)

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com os conceitos básicos do uso da API de diálogo do Office, conforme descrito em Usar a API de diálogo Office em seus [suplementos Office](dialog-api-in-office-add-ins.md).
> 
> Consulte também [Tratamento de erros e eventos com a Office caixa de diálogo.](dialog-handle-errors-events.md)

## <a name="rules-and-gotchas"></a>Regras e dicas

- A caixa de diálogo só pode navegar até URLs HTTPS, não HTTP.
- A URL passada para o [método displayDialogAsync](/javascript/api/office/office.ui) deve estar exatamente no mesmo domínio que o suplemento em si. Não pode ser um subdomínio. Mas a página que é passada para ela pode redirecionar para uma página em outro domínio.
- Uma janela de host, que pode ser um painel de tarefas ou o arquivo de função [](/javascript/api/manifest/functionfile) sem interface do usuário de um comando de suplemento, pode ter apenas uma caixa de diálogo aberta por vez.
- Somente duas Office APIs podem ser chamadas na caixa de diálogo:
  - A [função messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) .
  - `Office.context.requirements.isSetSupported`(Para obter mais informações, consulte [Especificar Office aplicativos e requisitos de API](specify-office-hosts-and-api-requirements.md).)
- A [função messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) geralmente deve ser chamada de uma página no mesmo domínio que o suplemento em si, mas isso não é obrigatório. Para obter mais informações, [mensagens entre domínios para o runtime do host](dialog-api-in-office-add-ins.md#cross-domain-messaging-to-the-host-runtime).

## <a name="best-practices"></a>Práticas recomendadas

### <a name="avoid-overusing-dialog-boxes"></a>Evitar o uso em desuso de caixas de diálogo

Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso. Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias. Para obter um exemplo de um painel de tarefas com guias, consulte [o exemplo Excel Suplemento JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

### <a name="design-a-dialog-box-ui"></a>Criar uma interface do usuário da caixa de diálogo

Para obter as práticas recomendadas no design da caixa de diálogo, consulte [caixas de diálogo Office suplementos](../develop/dialog-api-in-office-add-ins.md).

### <a name="handle-pop-up-blockers-with-office-on-the-web"></a>Manipular bloqueadores de pop-up com Office na Web

Tentar exibir uma caixa de diálogo ao Office na Web pode fazer com que o bloqueador pop-up do navegador bloqueie a caixa de diálogo. Se isso acontecer, Office na Web abrirá um prompt semelhante ao seguinte.

![Captura de tela mostrando o prompt com uma breve descrição e os botões Permitir e Ignorar que um suplemento pode gerar para evitar bloqueadores pop-up no navegador](../images/dialog-prompt-before-open.png)

Se o usuário escolher **Permitir, a** Office de diálogo será aberta. Se o usuário escolher **Ignorar**, o prompt será fechado e Office caixa de diálogo não será aberta. Em vez disso `displayDialogAsync` , o método retorna o erro 12009. Seu código deve detectar esse erro e fornecer uma experiência alternativa que não exija uma caixa de diálogo ou exibir uma mensagem para o usuário avisando que o suplemento exige que ele permita a caixa de diálogo. (Para obter mais informações sobre 12009, consulte [Erros de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)

Se, por qualquer motivo, você quiser desativar esse recurso, seu código deverá recusar. Ele faz essa solicitação com o [objeto DialogOptions](/javascript/api/office/office.dialogoptions) que é passado para o `displayDialogAsync` método. Especificamente, o objeto deve incluir `promptBeforeOpen: false`. Quando essa opção for definida como false, Office na Web solicitará que o usuário permita que o suplemento abra uma caixa de diálogo e a caixa de diálogo Office não será aberta.

### <a name="do-not-use-the-_host_info-value"></a>Não use o valor \_de hostinfo\_

Office adiciona automaticamente um parâmetro de consulta chamado `_host_info` à URL que é passada para `displayDialogAsync`. Ele é acrescentado após os parâmetros de consulta personalizados, se houver. Ele não é acrescentado a nenhuma URL subsequente para a qual a caixa de diálogo navega. A Microsoft pode alterar o conteúdo desse valor ou removê-lo inteiramente, portanto, seu código não deve lê-lo. O mesmo valor é adicionado ao armazenamento de sessão da caixa de diálogo (ou seja, a [propriedade Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). Novamente, *seu código não deve ler nem gravar nesse valor*.

### <a name="open-another-dialog-immediately-after-closing-one"></a>Abrir outra caixa de diálogo imediatamente após o fechamento de um

Você não pode ter mais de uma caixa de diálogo aberta em uma determinada página de host, portanto, seu código deve chamar [Dialog.close](/javascript/api/office/office.dialog#office-office-dialog-close-member(1)) `displayDialogAsync` em um diálogo aberto antes que ele chame para abrir outro diálogo. O `close` método é assíncrono. Por esse motivo, se `displayDialogAsync` você chamar imediatamente após uma chamada, o `close`primeiro diálogo poderá não ter sido completamente fechado quando Office tentar abrir o segundo. Se isso acontecer, Office retornará um erro [12007](dialog-handle-errors-events.md#12007): "A operação falhou porque esse suplemento já tem um diálogo ativo".

O `close` método não aceita um parâmetro de retorno de chamada e não retorna um objeto Promise `await` , portanto, ele não pode ser aguardado com a palavra-chave ou com um `then` método. Por esse motivo, sugerimos a seguinte técnica quando você precisar abrir uma nova caixa de diálogo imediatamente após fechar um diálogo: encapsular o código para abrir o novo diálogo em um método e projetar o método para chamar-se recursivamente `displayDialogAsync` se a chamada de retorna `12007`. Apresentamos um exemplo a seguir.

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

Como alternativa, você pode forçar o código a pausar antes que ele tente abrir a segunda caixa de diálogo usando o [método setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) . Apresentamos um exemplo a seguir.

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

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>Práticas recomendadas para usar a API Office caixa de diálogo em um SPA

Se o suplemento usa o roteamento do lado do cliente, como os SPAs (aplicativos de página única) normalmente usam, você tem a opção de passar a URL de uma rota para o método [displayDialogAsync](/javascript/api/office/office.ui) em vez da URL de uma página HTML separada. *É recomendável não fazer isso pelos motivos fornecidos abaixo.*

> [!NOTE]
> Este artigo não é relevante para o *roteamento do lado* do servidor, como em um aplicativo Web baseado em Express.

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>Problemas com SPAs e a API Office caixa de diálogo

A Office caixa de diálogo está em uma nova janela com sua própria instância do mecanismo JavaScript e, portanto, seu próprio contexto de execução completo. Se você passar uma rota, sua página base e todo o código de inicialização e inicialização serão executados novamente nesse novo contexto, e todas as variáveis serão definidas com seus valores iniciais na caixa de diálogo. Portanto, essa técnica baixa e inicia uma segunda instância do seu aplicativo na janela da caixa, o que derrota parcialmente a finalidade de um SPA. Além disso, o código que altera variáveis na janela da caixa de diálogo não altera a versão do painel de tarefas das mesmas variáveis. Da mesma forma, a janela da caixa de diálogo tem seu próprio armazenamento de sessão (a propriedade [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), que não está acessível do código no painel de tarefas. A caixa de diálogo e a página host na qual foi `displayDialogAsync` chamada se parecem com dois clientes diferentes para o servidor. (Para um lembrete do que é uma página host, consulte [Abrir uma caixa de diálogo de uma página host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).)

Portanto, se você passar `displayDialogAsync` uma rota para o método, não teria realmente um SPA; você teria duas *instâncias do mesmo SPA*. Além disso, grande parte do código na instância do painel de tarefas nunca seria usada nessa instância e grande parte do código na instância da caixa de diálogo nunca seria usada nessa instância. Seria como ter dois SPAs no mesmo grupo.

#### <a name="microsoft-recommendations"></a>Recomendações da Microsoft

Em vez de passar uma rota do lado do cliente para o `displayDialogAsync` método, recomendamos que você siga um dos seguintes procedimentos:

* Se o código que você deseja executar na caixa de diálogo for suficientemente complexo, crie dois SPAs diferentes explicitamente; ou seja, ter dois SPAs em pastas diferentes do mesmo domínio. Um SPA é executado na caixa de diálogo e o outro na página host da caixa de diálogo onde foi `displayDialogAsync` chamado. 
* Na maioria dos cenários, apenas a lógica simples é necessária na caixa de diálogo. Nesses casos, seu projeto será bastante simplificado hospedando uma única página HTML, com JavaScript inserido ou referenciado, no domínio do SPA. Passe a URL da página para o método`displayDialogAsync`. Embora isso significa que você está se desviando da ideia literal de um aplicativo de página única; você realmente não tem uma única instância de um SPA quando está usando a API de Office diálogo.
