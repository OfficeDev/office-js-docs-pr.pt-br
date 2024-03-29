---
title: Desenvolver suplementos do Office para o Angular
description: Use Angular para criar um Suplemento do Office como um aplicativo de página única.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: bbac0f94b731b2853e17ed3db785b50ea99ef6e4
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422953"
---
# <a name="develop-office-add-ins-with-angular"></a>Desenvolver suplementos do Office para o Angular

Este artigo fornece orientações sobre como usar o Angular 2+ para criar um Suplemento do Office como um aplicativo de página única.

> [!NOTE]
> Você tem alguma contribuição a fazer com base na sua experiência de uso do Angular para criar Suplementos do Office? Você pode contribuir para [este artigo no GitHub](https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/develop/add-ins-with-angular2.md) ou fornecer seus comentários enviando um [problema](https://github.com/OfficeDev/office-js-docs-pr/issues) no repositório.

Para ver um exemplo de Suplementos do Office criado utilizando a estrutura do Angular, confira o [Suplemento de Verificação de Estilo do Word Criado no Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="install-the-typescript-type-definitions"></a>Instalar as definições de tipo TypeScript

Abra uma Node.js e insira o seguinte na linha de comando.

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>A inicialização deve ocorrer no Office.initialize

Em qualquer página que chame as APIs JavaScript do Office, word ou Excel, seu código deve primeiro atribuir uma função `Office.initialize`a . (Se você não tiver nenhum código de inicialização, o corpo da função poderá ser apenas símbolos "`{}`" vazios, mas você não deverá deixar `Office.initialize` a função indefinida. Para obter detalhes, [consulte Inicializar seu Suplemento do Office](initialize-add-in.md).) O Office chama essa função imediatamente depois de inicializar as bibliotecas JavaScript do Office.

**Seu Angular código de inicialização `Office.initialize`** deve ser chamado dentro da função que você atribui para garantir que as bibliotecas JavaScript do Office tenham sido inicializadas primeiro. O exemplo a seguir mostra como fazer isso. Este código deve estar no arquivo main.ts do projeto.

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>Use a estratégia de localização de hash no aplicativo Angular

A navegação entre rotas no aplicativo pode não funcionar se você não especificar a estratégia de localização de hash. Você pode fazer isso de duas maneiras. Primeiro, você pode especificar um provedor para a estratégia de localização no módulo do seu aplicativo, conforme mostrado no exemplo a seguir. Ele fica no arquivo app.module.ts.

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity

@NgModule({
  providers: [
    { provide: LocationStrategy, useClass: HashLocationStrategy },
    // Other providers suppressed
  ],
  // Other module properties suppressed
})
export class AppModule { }
```

Se você definir suas rotas em um módulo de roteamento distinto, há uma forma alternativa para especificar a estratégia de localização de hash. No seu arquivo .ts do módulo de roteamento, transmita um objeto de configuração para a função `forRoot` que especifica a estratégia. O código a seguir é um exemplo.

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity

const routes: Routes = // route definitions go here

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
```

## <a name="use-the-office-dialog-api-with-angular"></a>Usar a API de caixa de diálogo do Office com Angular

A API da caixa de diálogo Suplemento do Office permite que o suplemento abra uma página em uma caixa de diálogo não modal que pode trocar informações com a página principal, normalmente em um painel de tarefas.

O método [displayDialogAsync](/javascript/api/office/office.ui) usa um parâmetro que especifica a URL da página que deve ser aberta na caixa de diálogo. Seu suplemento pode ter uma página HTML distinta (diferente da página de base) para transmitir esse parâmetro, ou você pode transmitir a URL de uma rota em um aplicativo do Angular.

É importante lembrar, se você transmitir uma rota, que a caixa de diálogo cria uma nova janela com seu próprio contexto de execução. Sua página de base e todos os códigos de inicialização são executados novamente nesse novo contexto e todas as variáveis são definidas para seus valores iniciais na caixa de diálogo. Então essa técnica lança uma segunda instância do seu aplicativo com uma única página na caixa de diálogo. O código que altera as variáveis na caixa de diálogo não altera a versão do painel tarefas das mesmas variáveis. Da mesma forma, a caixa de diálogo tem seu próprio armazenamento de sessão (a [propriedade Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), que não está acessível do código no painel de tarefas.  

## <a name="trigger-the-ui-update"></a>Acionar a atualização da interface do usuário

Em um aplicativo do Angular, às vezes, a interface do usuário não é atualizada. Isso ocorre porque essa parte do código fica sem a zona do Angular. A solução é colocar o código na região, conforme mostrado no exemplo a seguir.

```js
import { NgZone } from '@angular/core';

export class MyComponent {
  constructor(private zone: NgZone) { }

  myFunction() {
    this.zone.run(() => {
      // the codes that need update the UI
    });
  }
}
```

## <a name="use-observable"></a>Usar Observável

O Angular usa o RxJS (Expansões Reativas para JavaScript) e o RxJS introduz os objetos `Observable` e `Observer` para implementar o processamento assíncrono. Esta seção fornece uma breve introdução ao uso de `Observables`; para saber mais informações, consulte a documentação de [RxJS](https://rxjs-dev.firebaseapp.com/) oficial.

Um `Observable` é como um objeto `Promise` em certos aspectos. Ele é retornado diretamente de uma chamada assíncrona, mas poderá só ser resolvido algum tempo depois. Contudo, embora `Promise` seja um único valor (que pode ser um objeto de matriz), um `Observable` é uma matriz de objetos (possivelmente com apenas um único membro). Isso permite que o código chame [métodos de matriz](https://www.w3schools.com/jsref/jsref_obj_array.asp), como `concat`, `map` e `filter`, em objetos `Observable`.

### <a name="push-instead-of-pull"></a>Enviar por push em vez de efetuar pull

Seu código "obtém" objetos `Promise` atribuindo-os a variáveis, mas objetos `Observable` "enviam" seus valores para objetos que se *inscrevem* no `Observable`. Os assinantes são objetos `Observer`. O benefício da arquitetura push é que novos membros podem ser adicionados à matriz `Observable` ao longo do tempo. Quando um novo membro é adicionado, todos os objetos `Observer` que assinam o `Observable` recebem uma notificação.

O `Observer` é configurado para processar cada novo objeto (chamado o "próximo" objeto) com uma função. (Ele também é configurado para responder a um erro e a uma notificação de conclusão. Consulte a próxima seção para obter um exemplo.) Por esse motivo, os objetos `Observable` podem ser usados em uma maior variedade de cenários do que os objetos `Promise`. Por exemplo, além de retornarem um `Observable` de uma chamada AJAX, a maneira como você pode retornar um `Promise`, um `Observable` pode ser retornado de um manipulador de eventos, como o manipulador de eventos "modificado" de uma caixa de texto. Cada vez que um usuário insere texto na caixa, todos os objetos `Observer` inscritos reagem imediatamente usando o texto mais recente e/ou o estado atual do aplicativo como entrada.

### <a name="wait-until-all-asynchronous-calls-have-completed"></a>Aguarde até que todas as chamadas assíncronas sejam concluídas

Quando quiser garantir que um retorno de chamada só seja executado quando todos os membros de um conjunto de objetos `Promise` forem resolvidos, use o método `Promise.all()`.

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
```

Para fazer o mesmo com um objeto `Observable`, use o método [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).  

```js
const source = Observable.forkJoin([x, y, z]);

const subscription = source.subscribe(
  x => {
    // TODO: Callback logic goes here
  },
  err => console.log('Error: ' + err),
  () => console.log('Completed')
);
```

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a>Compilar o aplicativo Angular usando o compilador AOT (Ahead-of-Time)

O desempenho do aplicativo é um dos aspectos mais importantes da experiência do usuário. Um aplicativo Angular pode ser otimizado usando o compilador Angular AOT (Ahead-of-Time) para compilar o aplicativo durante a compilação. Ele converte todo o código-fonte (modelos HTML e TypeScript) em um código JavaScript eficiente. Se você compilar o aplicativo com o compilador AOT, nenhuma compilação adicional ocorrerá no tempo de execução, o que resultará em um processamento mais rápido e solicitações assíncronas mais rápidas para modelos HTML. Além disso, o tamanho geral do aplicativo diminui, pois o compilador Angular não precisa ser incluído no aplicativo para distribuição.

Para usar o compilador AOT, adicione `--aot` aos comandos `ng build` ou `ng serve`:

```command&nbsp;line
ng build --aot
ng serve --aot
```

> [!NOTE]
> Para saber mais sobre o compilador Angular AOT (Ahead-of-Time), consulte o [guia oficial](https://angular.io/guide/aot-compiler).

## <a name="support-internet-explorer-if-youre-dynamically-loading-officejs"></a>Dar suporte ao Internet Explorer se você estiver carregando dinamicamente Office.js

Com base na versão do Windows e no cliente da área de trabalho do Office em que o suplemento está em execução, o suplemento pode estar usando o Internet Explorer 11. (Para obter mais detalhes, consulte [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).) Angular depende de alguns `window.history` APIs, mas essas APIs não funcionam no runtime do IE que às vezes é usado para executar Suplementos do Office em clientes da área de trabalho do Windows. Quando essas APIs não funcionam, seu suplemento pode não funcionar corretamente, por exemplo, ele pode carregar um painel de tarefas em branco. Para atenuar isso, Office.js anula essas APIs. No entanto, se você estiver carregando dinamicamente Office.js, o AngularJS poderá ser carregado antes Office.js. Nesse caso, você deve desabilitar as `window.history` APIs adicionando o código a seguir à página de **index.htmldo suplemento** .

```js
<script type="text/javascript">window.history.replaceState=null;window.history.pushState=null;</script>
```
