---
title: Runtimes em Suplementos do Office
description: Saiba mais sobre os runtimes usados pelos Suplementos do Office.
ms.date: 08/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8d28f6db028d2f4c7036db51ccc5dbcc2144bdf3
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616039"
---
# <a name="runtimes-in-office-add-ins"></a>Runtimes em Suplementos do Office

Os Suplementos do Office são executados em runtimes inseridos no Office. Como uma linguagem interpretada, o JavaScript deve ser executado em um mecanismo JavaScript. Como uma linguagem síncrona de thread único, o JavaScript não tem capacidade inerente para execução simultânea; mas os mecanismos modernos do JavaScript podem solicitar operações simultâneas (incluindo comunicação de rede) do sistema operacional host e receber dados do sistema operacional em resposta. Esse tipo de mecanismo torna o JavaScript *efetivamente* assíncrono. Neste artigo, os mecanismos desse tipo são chamados *de runtimes*. [Node.js](https://nodejs.org) e navegadores modernos são exemplos desses runtimes. 

## <a name="types-of-runtimes"></a>Tipos de runtimes

Há dois tipos de runtimes usados pelos Suplementos do Office:

- **Runtime somente JavaScript**: um mecanismo JavaScript complementado com suporte para [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [CORS Completo (](https://developer.mozilla.org/docs/Web/HTTP/CORS)Compartilhamento de Recursos entre Origens) e armazenamento de dados do lado do cliente. (Ele não dá suporte ao [armazenamento local ou](https://developer.mozilla.org/docs/Web/API/Window/localStorage) cookies.) 
- **Runtime do navegador**: inclui todos os recursos de um runtime somente JavaScript e adiciona suporte para armazenamento [local](https://developer.mozilla.org/docs/Web/API/Window/localStorage)[, mecanismo](https://developer.mozilla.org/docs/Glossary/Rendering_engine) de renderização que renderiza HTML e cookies.

Os detalhes sobre esses tipos são posteriormente neste artigo no [runtime somente javaScript](#javascript-only-runtime) e [no runtime do navegador](#browser-runtime).

A tabela a seguir mostra quais recursos possíveis de um suplemento usam cada tipo de runtime. 

> [!NOTE]
> A escolha de qual tipo de runtime usar é um detalhe de implementação que a Microsoft pode alterar a qualquer momento. A Biblioteca JavaScript do Office não pressupõe que o mesmo tipo de runtime sempre será usado para um determinado recurso e sua arquitetura de suplemento também não deve pressupor isso.

| Tipo de runtime | Recurso de suplemento |
|:-----|:-----|
| Somente JavaScript | Funções [personalizadas do](../excel/custom-functions-overview.md) Excel</br>(exceto quando o runtime é [compartilhado](#shared-runtime) ou o suplemento está em execução no Office na Web)</br></br>[Tarefa baseada em eventos do Outlook](../outlook/autolaunch.md)</br>(somente quando o suplemento estiver em execução no Outlook no Windows)|
| Navegador | [painel de tarefas](../design/task-pane-add-ins.md)</br></br>[caixa de diálogo](../develop/dialog-api-in-office-add-ins.md)</br></br>[comando de função](../design/add-in-commands.md#types-of-add-in-commands)</br></br>Funções [personalizadas do](../excel/custom-functions-overview.md) Excel</br>(quando o runtime é [compartilhado](#shared-runtime) ou o suplemento está em execução no Office na Web)</br></br>[Tarefa baseada em eventos do Outlook](../outlook/autolaunch.md)</br>(quando o suplemento está em execução no Outlook no Mac ou Outlook na Web)|

A tabela a seguir mostra as mesmas informações organizadas por qual tipo de runtime é usado para os vários recursos possíveis de um suplemento.

| Recurso de suplemento | Tipo de runtime no Windows | Tipo de runtime no Mac | Tipo de runtime na Web |
|:-----|:-----|:-----|:-----|
|Funções personalizadas do Excel | Somente JavaScript</br>(mas *navegador* quando o runtime é compartilhado)|Somente JavaScript</br>(mas *navegador* quando o runtime é compartilhado)| Navegador |
|Tarefas baseadas em eventos do Outlook | Somente JavaScript | Navegador | Navegador |
|painel de tarefas | Navegador | Navegador | Navegador |
|caixa de diálogo | Navegador | Navegador | Navegador |
|comando de função | Navegador | Navegador | Navegador |


No Office na Web, tudo sempre é executado em um runtime de tipo de navegador. Na verdade, com uma exceção, tudo em um suplemento na Web é executado no mesmo processo  do navegador: o processo do navegador no qual o usuário abriu Office na Web. A exceção é quando uma caixa de diálogo é aberta com uma chamada de [Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) e a [opção DialogOptions.displayInIFrame](/javascript/api/office/office.dialogoptions#office-office-dialogoptions-displayiniframe-member) não é passada e definida como . `true` Quando a opção não é passada (portanto, ela tem o valor `false` padrão), a caixa de diálogo é aberta em seu próprio processo. O mesmo princípio se aplica ao método [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) e à opção [OfficeRuntime.DisplayWebDialogOptions.displayInIFrame](/javascript/api/office-runtime/officeruntime.displaywebdialogoptions#office-runtime-officeruntime-displaywebdialogoptions-displayiniframe-member) .

Quando um suplemento está em execução em uma plataforma diferente da Web, os princípios a seguir se aplicam.

- Uma caixa de diálogo é executada em seu próprio processo de runtime. 
- Uma tarefa baseada em evento do Outlook é executada em seu próprio processo de runtime. 
- Por padrão, painéis de tarefas, comandos de função e funções personalizadas do Excel são executados em seu próprio processo de runtime. No entanto, para alguns aplicativos host do Office, o manifesto do suplemento pode ser configurado para que qualquer um dos dois ou todos os três possa ser executado no mesmo runtime. Consulte [Runtime compartilhado](#shared-runtime).

Dependendo do aplicativo host do Office e dos recursos usados no suplemento, pode haver muitos runtimes em um suplemento. Cada um geralmente será executado em seu próprio processo, mas não necessariamente simultaneamente. Eis alguns exemplos.

- Um suplemento do PowerPoint ou do Word que não compartilha nenhum runtime e inclui os recursos a seguir tem até três runtimes.

  - Um painel de tarefas
  - Um comando de função
  - Uma caixa de diálogo (uma caixa de diálogo pode ser iniciada no painel de tarefas ou no comando de função.) 
  
      > [!NOTE]
      > Não é uma boa prática ter vários diálogos abertos simultaneamente, mas se o suplemento permitir que o usuário abra um no painel de tarefas e outro no comando de função ao mesmo tempo, esse suplemento terá quatro runtimes. Um painel de tarefas e uma determinada invocação de um comando de função podem ter apenas um diálogo aberto por vez; mas se o comando de função for invocado várias vezes, uma nova caixa de diálogo será aberta sobre seu predecessor com cada invocação, portanto, pode haver muitos runtimes. O restante desta lista ignora a possibilidade de vários diálogos abertos.

- Um suplemento do Excel que não compartilha nenhum runtime e inclui os recursos a seguir tem até *quatro* runtimes.

  - Um painel de tarefas
  - Um comando de função
  - Uma função personalizada
  - Uma caixa de diálogo (uma caixa de diálogo pode ser iniciada no painel de tarefas, no comando de função ou em uma função personalizada.)

- Um suplemento do Excel com os mesmos recursos e configurado para compartilhar o mesmo runtime no painel de tarefas, no comando de função e na função personalizada tem *dois runtimes* . Um runtime compartilhado pode abrir apenas uma caixa de diálogo por vez.
- Um suplemento do Excel com os mesmos recursos, exceto que não tem caixa de diálogo e está configurado para compartilhar o mesmo runtime no painel de tarefas, no comando de função e na função personalizada, tem *um runtime* .
- Um suplemento do Outlook que tem os seguintes recursos tem até *quatro* runtimes. (Os runtimes não podem ser compartilhados no Outlook.)

  - Um painel de tarefas
  - Um comando de função
  - Uma tarefa baseada em evento
  - Uma caixa de diálogo (uma caixa de diálogo pode ser iniciada no painel de tarefas ou no comando da função, mas não de uma tarefa baseada em evento.)

## <a name="share-data-across-runtimes"></a>Compartilhar dados entre runtimes

> [!NOTE]
> - Se você souber que `displayInIFrame` `true`seu suplemento será usado apenas no Office na Web e que ele não abrirá nenhuma caixa de diálogo com a opção definida como, você poderá ignorar esta seção. Como tudo em seu suplemento é executado no mesmo processo de runtime, você pode usar apenas variáveis globais para compartilhar dados entre recursos.
> - Conforme mencionado acima em [Tipos de runtimes](#types-of-runtimes), o tipo de runtime usado por um recurso varia parcialmente por plataforma. É uma boa prática evitar ter código de suplemento que se ramifica com base na plataforma, portanto, as diretrizes nesta seção recomendam técnicas que funcionarão em plataforma cruzada. Há apenas um caso, descrito abaixo, no qual o código de ramificação é necessário. 

Para suplementos do Excel, do PowerPoint e do Word, use um [runtime](#shared-runtime) compartilhado quando dois ou mais recursos, exceto caixas de diálogo, precisam compartilhar dados. No Outlook ou em cenários em que o compartilhamento de um runtime não é viável, você precisa de métodos alternativos. As partes do suplemento que estão em processos de runtime separados não compartilham dados globais automaticamente e são tratadas pelo servidor de aplicativos Web do suplemento como sessões separadas, portanto [, Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) não pode ser usado para compartilhar dados entre eles. *As diretrizes a seguir pressupõem que você não está usando um runtime compartilhado.*

- Passe dados entre uma caixa de diálogo e seu painel de tarefas pai, comando de função ou função personalizada usando os métodos [Office.ui.messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) e [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) . 

    > [!NOTE]
    > Os `OfficeRuntime.storage` métodos não podem ser chamados em uma caixa de diálogo, portanto, essa não é uma opção para compartilhar dados entre uma caixa de diálogo e outro runtime. 

- Para compartilhar dados entre um painel de tarefas e um comando de função, armazene dados em [Window.localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), que são compartilhados entre todos os runtimes que acessam a mesma [origem específica](https://developer.mozilla.org/docs/Glossary/Origin). 
    > [!NOTE]
    > O LocalStorage não está acessível em um runtime somente JavaScript e, portanto, não está disponível em funções personalizadas do Excel. Ele também não pode ser usado para compartilhar dados com tarefas baseadas em eventos do Outlook (já que essas tarefas usam um runtime somente JavaScript em algumas plataformas).

    > [!TIP]
    > Os dados `Window.localStorage` em persistem entre as sessões do suplemento e são compartilhados por suplementos com a mesma origem. Essas duas características geralmente são indesejáveis para um suplemento. 
    >
    > - Para garantir que cada sessão de um determinado suplemento comece a chamar o método [Window.localStorage.clear](https://developer.mozilla.org/docs/Web/API/Storage/clear) quando o suplemento for iniciado. 
    > - Para permitir que alguns valores armazenados persistam, mas reinicializar outros valores, use [Window.localStorage.setItem](https://developer.mozilla.org/docs/Web/API/Storage/setItem) quando o suplemento for iniciado para cada item que deve ser redefinido para um valor inicial. 
    > - Para excluir um item inteiramente, chame [Window.localStorage.removeItem](https://developer.mozilla.org/docs/Web/API/Storage/removeItem).

- Para compartilhar dados entre uma função personalizada do Excel e qualquer outro runtime, use [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage).
- Para compartilhar dados entre uma tarefa baseada em evento do Outlook e um painel de tarefas ou comando de função, você deve ramificar seu código pelo valor da propriedade [Office.context.platform](/javascript/api/office/office.context#office-office-context-platform-member) . 

    - Quando o valor for `PC` (Windows), armazene e recupere dados usando as APIs [Office.sessionData](/javascript/api/outlook/office.sessiondata) .
    - Quando o valor for `Mac`, use `Window.localStorage` conforme descrito anteriormente nesta lista.

Outras maneiras de compartilhar dados incluem o seguinte:

- Armazene dados compartilhados em um banco de dados online acessível a todos os runtimes.
- Armazene dados compartilhados em um cookie para o domínio do suplemento compartilhá-los entre runtimes do navegador. Os runtimes somente para JavaScript não dão suporte a cookies.

Para obter mais informações, [consulte Persistir](../develop/persisting-add-in-state-and-settings.md) o estado e as configurações do suplemento e gerenciar o estado e as configurações [de um suplemento do Outlook](../outlook/manage-state-and-settings-outlook.md).

## <a name="javascript-only-runtime"></a>Runtime somente JavaScript

O runtime somente JavaScript usado em Suplementos do Office é uma modificação de um runtime código aberto criado originalmente para [React Native](https://reactnative.dev/). Ele contém um mecanismo JavaScript complementado com suporte para [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [CORS Completo (](https://developer.mozilla.org/docs/Web/HTTP/CORS)Compartilhamento de Recursos entre Origens) e [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage). Ele não tem um mecanismo de renderização e não dá suporte a cookies ou [armazenamento local](https://developer.mozilla.org/docs/Web/API/Window/localStorage).

Esse tipo de runtime é usado em tarefas baseadas em eventos do Outlook somente no Office no Windows e em funções *personalizadas do* Excel, exceto quando as funções personalizadas estão compartilhando [um runtime](#shared-runtime). 

- Quando usado para uma função personalizada do Excel, o runtime é iniciado quando a planilha é recalculada ou a função personalizada é calculada. Ele não é desligado até que a pasta de trabalho seja fechada.  
- Quando usado em uma tarefa baseada em evento do Outlook, o runtime é iniciado quando o evento ocorre. Ele termina quando ocorre o primeiro dos itens a seguir.

  - O manipulador de eventos chama o `completed` método de seu parâmetro de evento.
  - 5 minutos decorridos desde o evento de gatilho.
  - O usuário altera o foco da janela em que o evento foi disparado, como uma janela de composição de mensagem.

Um javaScript-runtime usa menos memória e inicia mais rapidamente do que um runtime do navegador, mas tem menos recursos.

## <a name="browser-runtime"></a>Runtime do navegador

Os Suplementos do Office usam um runtime de tipo de navegador diferente, dependendo da plataforma na qual o Office está em execução (Web, Mac ou Windows) e na versão e build do Windows e do Office. Por exemplo, se o usuário estiver executando Office na Web em um navegador FireFox, o runtime do Firefox será usado. Se o usuário estiver executando o Office no Mac, o runtime do Safari será usado. Se o usuário estiver executando o Office no Windows, um Edge ou o Internet Explorer fornecerá o runtime, dependendo da versão do Windows e do Office. Os detalhes podem ser [encontrados em Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

Todos esses runtimes incluem um mecanismo de renderização HTML e fornecem suporte para [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [CORS Completo (](https://developer.mozilla.org/docs/Web/HTTP/CORS)Compartilhamento de Recursos entre Origens) e armazenamento [local](https://developer.mozilla.org/docs/Web/API/Window/localStorage) e cookies. 

Um tempo de vida de runtime do navegador varia dependendo do recurso que ele implementa e se ele está sendo compartilhado ou não.

- Quando um suplemento com um painel de tarefas é iniciado, um runtime do navegador é iniciado, a menos que seja um runtime compartilhado que já esteja em execução. Se for um runtime compartilhado, ele será desligado quando o documento for fechado. Se não for um runtime compartilhado, ele será desligado quando o painel de tarefas for fechado.
- Quando uma caixa de diálogo é aberta, um runtime do navegador é iniciado. Ele é desligado quando a caixa de diálogo é fechada.
- Quando um comando de função é executado (o que acontece quando um usuário seleciona seu botão ou item de menu), um runtime do navegador é iniciado, a menos que seja um runtime compartilhado que já esteja em execução. Se for um runtime compartilhado, ele será desligado quando o documento for fechado. Se não for um runtime compartilhado, ele será desligado quando o primeiro dos seguintes eventos ocorrer.
 
  - O comando de função chama o `completed` método de seu parâmetro de evento.
  - 5 minutos decorridos desde o evento de gatilho. (Se uma caixa de diálogo tiver sido aberta no comando de função e ainda estiver aberta quando o runtime pai expirar, o runtime do diálogo permanecerá em execução até que a caixa de diálogo seja fechada.)

- Quando uma função personalizada do Excel está usando um runtime compartilhado, um runtime do tipo navegador é iniciado quando a função personalizada calcula se o runtime compartilhado ainda não foi iniciado por algum outro motivo. Ele é desligado quando o documento é fechado.

> [!NOTE]
> Quando um runtime está sendo [compartilhado](#shared-runtime), é possível que seu código feche o painel de tarefas sem desligar o suplemento. Consulte [Mostrar ou ocultar o painel de tarefas do suplemento do Office](../develop/show-hide-add-in.md) para obter mais informações.

Um runtime do navegador tem mais recursos do que um runtime somente JavaScript, mas é iniciado mais lentamente e usa mais memória.

### <a name="shared-runtime"></a>Tempo de execução compartilhado

Um "runtime compartilhado" não é um tipo de runtime. Ele se refere a um [runtime](#browser-runtime) do tipo navegador que está sendo compartilhado por recursos do suplemento que, caso contrário, cada um teria seu próprio runtime. Especificamente, você tem a opção de configurar o painel de tarefas do suplemento e os comandos de função para compartilhar um runtime. Em um suplemento do Excel, você também pode configurar funções personalizadas para compartilhar o runtime de um painel de tarefas, comando de função ou ambos. Quando você faz isso, as funções personalizadas são executadas em um runtime do tipo navegador, em vez de um [runtime somente JavaScript](#javascript-only-runtime) como faria de outra forma. Consulte Configurar seu suplemento para usar um [runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) compartilhado para obter informações sobre os benefícios e limitações do compartilhamento de runtimes e instruções para configurar o suplemento para usar um runtime compartilhado. Em resumo, o runtime somente JavaScript usa menos memória e inicia mais rapidamente, mas tem menos recursos.

> [!NOTE]
> - Você pode compartilhar runtimes somente no Excel, no PowerPoint e no Word. 
> - Não é possível configurar uma caixa de diálogo para compartilhar um runtime. Cada caixa de diálogo sempre tem sua própria, exceto quando a caixa de diálogo é iniciada Office na Web com `displayInIFrame` a opção definida como `true`.
> - Um runtime compartilhado nunca usa o runtime original do Microsoft Edge WebView (EdgeHTML). Se as condições para usar o Microsoft Edge com o WebView2 (baseado em Chromium) forem atendidas (conforme especificado em Navegadores usados pelos [Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md)), esse runtime será usado. Caso contrário, o runtime do Internet Explorer 11 será usado.