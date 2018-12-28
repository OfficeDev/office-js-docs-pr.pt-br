---
title: Desenvolver suplementos do Office para o Angular
description: ''
ms.date: 11/02/2018
ms.openlocfilehash: 0ae27efb3a89244e76860048d6c8ef4d78b613d1
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457856"
---
# <a name="develop-office-add-ins-with-angular"></a><span data-ttu-id="e53dc-102">Desenvolver suplementos do Office para o Angular</span><span class="sxs-lookup"><span data-stu-id="e53dc-102">Develop Office Add-ins with Angular</span></span>

<span data-ttu-id="e53dc-103">Este artigo fornece orientações sobre como usar o Angular 2+ para criar um Suplemento do Office como um aplicativo de página única.</span><span class="sxs-lookup"><span data-stu-id="e53dc-103">This article provides guidance for using Angular 2+ to create an Office Add-in as a single page application.</span></span>

> [!NOTE]
> <span data-ttu-id="e53dc-p101">Você tem alguma contribuição a fazer com base na sua experiência de uso do Angular para criar Suplementos do Office? É possível contribuir com este artigo no [GitHub](https://github.com/OfficeDev/office-js-docs) ou enviando seus comentários por meio de uma [questão](https://github.com/OfficeDev/office-js-docs-pr/issues) no repositório.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p101">Do you have something to contribute based on your experience using Angular to create Office Add-ins? You can contribute to this article in [GitHub](https://github.com/OfficeDev/office-js-docs) or provide your feedback by submitting an [issue](https://github.com/OfficeDev/office-js-docs-pr/issues) in the repo.</span></span> 

<span data-ttu-id="e53dc-106">Para ver um exemplo de Suplementos do Office criado utilizando a estrutura do Angular, confira o [Suplemento de Verificação de Estilo do Word Criado no Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span><span class="sxs-lookup"><span data-stu-id="e53dc-106">For an Office Add-ins sample that's built using the Angular framework, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span></span>

## <a name="install-the-typescript-type-definitions"></a><span data-ttu-id="e53dc-107">Instalar as definições de tipo TypeScript</span><span class="sxs-lookup"><span data-stu-id="e53dc-107">Install the TypeScript type definitions</span></span>
<span data-ttu-id="e53dc-108">Abra uma janela de nodejs e insira o seguinte na linha de comando:</span><span class="sxs-lookup"><span data-stu-id="e53dc-108">Open an nodejs window and enter the following at the command line:</span></span> 

```bash
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a><span data-ttu-id="e53dc-109">A inicialização deve ocorrer no Office.initialize</span><span class="sxs-lookup"><span data-stu-id="e53dc-109">Bootstrapping must be inside Office.initialize</span></span>

<span data-ttu-id="e53dc-p102">Em qualquer página que chame as APIs do JavaScript do Office, do Word ou do Excel, seu código deve atribuir primeiro um método para a propriedade `Office.initialize`. (Se você não tiver nenhum código de inicialização, o corpo do método poderá ser apenas símbolos "`{}`" vazios, mas você não deve deixar a propriedade `Office.initialize` indefinida. Para mais detalhes, confira [Iniciar o suplemento](understanding-the-javascript-api-for-office.md#initializing-your-add-in).) O Office chama esse método imediatamente depois que ele inicializa as bibliotecas JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p102">On any page that calls the Office, Word, or Excel JavaScript APIs, your code must first assign a method to the `Office.initialize` property. (If you have no initialization code, the method body can be just empty "`{}`" symbols, but you must not leave the `Office.initialize` property undefined. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).) Office calls this method immediately after it has initialized the Office JavaScript libraries.</span></span>

<span data-ttu-id="e53dc-p103">**O seu código de inicialização do Angular deve ser chamado dentro do método atribuído a `Office.initialize`** para garantir que as bibliotecas JavaScript do Office inicializem primeiro. O exemplo a seguir mostra como fazer isso. Este código deve estar no arquivo main.ts do projeto.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p103">**Your Angular bootstrapping code must be called inside the method that you assign to `Office.initialize`** to ensure that the Office JavaScript libraries have initialized first. The following is a simple example that shows how to do this. This code should be in the main.ts file of the project.</span></span>

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a><span data-ttu-id="e53dc-116">Use a estratégia de localização de hash no aplicativo Angular</span><span class="sxs-lookup"><span data-stu-id="e53dc-116">Use the hash location strategy in the Angular application</span></span>

<span data-ttu-id="e53dc-p104">A navegação entre rotas no aplicativo pode não funcionar se você não especificar a estratégia de localização de hash. Você pode fazer isso de duas maneiras. Primeiro, você pode especificar um provedor para a estratégia de localização no módulo do seu aplicativo, conforme mostrado no exemplo a seguir. Ele fica no arquivo app.module.ts.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p104">Navigating between routes in the application might not work if you don't specify the hash location strategy. You can do this in one of two ways. First, you can specify a provider for the location strategy in your app module, as shown in the following example. It goes into the app.module.ts file.</span></span>

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

<span data-ttu-id="e53dc-p105">Se você definir suas rotas em um módulo de roteamento distinto, há uma forma alternativa para especificar a estratégia de localização de hash. No seu arquivo .ts do módulo de roteamento, transmita um objeto de configuração para a função `forRoot` que especifica a estratégia. O código a seguir é um exemplo.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p105">If you define your routes in a separate routing module, there is an alternative way to specify the hash location strategy. In your routing module's .ts file, pass a configuration object to the `forRoot` function that specifies the strategy. The following code is an example.</span></span> 

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


## <a name="consider-wrapping-fabric-components-with-angular-components"></a><span data-ttu-id="e53dc-124">Considere a possibilidade de dispor componentes do Fabric com componentes do Angular</span><span class="sxs-lookup"><span data-stu-id="e53dc-124">Consider wrapping Fabric components with Angular components</span></span>

<span data-ttu-id="e53dc-p106">Recomendamos usar o uso do estilo [Office UI Fabric](https://developer.microsoft.com/fabric#/fabric-js) no seu suplemento. O Fabric contém componentes que vêm com em várias versões, incluindo uma versão [baseada no TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Considere o uso de componentes do Fabric no seu suplemento dispondo-os em componentes do Angular. Para ver um exemplo de como fazer isso, consulte [Suplemento de verificação de estilo do Word criado no Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Observe, por exemplo, como o componente do Angular definido em [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) importa o arquivo do Fabric TextField.ts, onde o componente do Fabric é definido.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p106">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric#/fabric-js) styling in your add-in. Fabric includes components that come in several versions, including a version [based on TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Consider using Fabric components in your add-in by wrapping them in Angular components. For an example that shows you how to do this, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Note, for example, how the Angular component defined in [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) imports the Fabric file TextField.ts, where the Fabric component is defined.</span></span> 


## <a name="using-the-office-dialog-api-with-angular"></a><span data-ttu-id="e53dc-130">Usar a API de caixa diálogo do Office com o Angular</span><span class="sxs-lookup"><span data-stu-id="e53dc-130">Using the Office Dialog API with Angular</span></span>

<span data-ttu-id="e53dc-131">A API de caixa de diálogo do Suplemento do Office permite que seu suplemento abra uma página em uma caixa de diálogo semimodal que pode trocar informações com a página principal, que, em geral, está no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="e53dc-131">The Office Add-in Dialog API enables your add-in to open a page in a semimodal dialog box that can exchange information with the main page, which is typically in a task pane.</span></span> 

<span data-ttu-id="e53dc-p107">O método [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui) usa um parâmetro que especifica a URL da página que deve ser aberta na caixa de diálogo. Seu suplemento pode ter uma página HTML distinta (diferente da página de base) para transmitir esse parâmetro ou você pode transmitir a URL de uma rota em um aplicativo do Angular.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p107">The [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui) method takes a parameter that specifies the URL of the page that should open in the dialog box. Your add-in can have a separate HTML page (different from the base page) to pass to this parameter, or you can pass the URL of a route in your Angular appication.</span></span> 

<span data-ttu-id="e53dc-p108">É importante lembrar, se você transmitir uma rota, que a caixa de diálogo cria uma nova janela com seu próprio contexto de execução. Sua página de base e todos os códigos de inicialização são executados novamente nesse novo contexto e todas as variáveis são definidas para seus valores iniciais na caixa de diálogo. Então essa técnica lança uma segunda instância do seu aplicativo com uma única página na caixa de diálogo. O código que altera as variáveis na caixa de diálogo não altera a versão do painel tarefas das mesmas variáveis. Da mesma forma, a caixa de diálogo tem seu próprio armazenamento de sessão, que não pode ser acessado a partir do código no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p108">It is important to remember, if you pass a route, that the dialog box creates a new window with its own execution context. Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box. So this technique launches a second instance of your single page application in the dialog box. Code that changes variables in the dialog box does not change the task pane version of the same variables. Similarly, the dialog box has its own session storage, which is not accessible from code in the task pane.</span></span>  


## <a name="trigger-the-ui-update"></a><span data-ttu-id="e53dc-139">Acionar a atualização da interface do usuário</span><span class="sxs-lookup"><span data-stu-id="e53dc-139">Trigger the UI update</span></span>

<span data-ttu-id="e53dc-p109">Em um aplicativo do Angular, às vezes, a interface do usuário não é atualizada. Isso ocorre porque essa parte do código fica sem a zona do Angular. A solução é colocar o código na região, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p109">In an Angular app, the UI sometimes does not update. This is because that part of the code runs out of the Angular zone. The solution is to put the code in the zone, as shown in the following example.</span></span>

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

## <a name="using-observable"></a><span data-ttu-id="e53dc-143">Usando o Observable</span><span class="sxs-lookup"><span data-stu-id="e53dc-143">Using Observable</span></span>

<span data-ttu-id="e53dc-p110">O Angular usa o RxJS (Expansões Reativas para JavaScript) e o RxJS introduz os objetos `Observable` e `Observer` para implementar o processamento assíncrono. Esta seção fornece uma breve introdução ao uso de `Observables`; para saber mais informações, consulte a documentação de [RxJS](https://rxjs-dev.firebaseapp.com/) oficial.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p110">Angular uses RxJS (Reactive Extensions for JavaScript), and RxJS introduces `Observable` and `Observer` objects to implement asynchronous processing. This section provides a brief introduction to using `Observables`; for more detailed information, see the official [RxJS](https://rxjs-dev.firebaseapp.com/) documentation.</span></span>

<span data-ttu-id="e53dc-p111">Um `Observable` é como um objeto `Promise` em certos aspectos. Ele é retornado diretamente de uma chamada assíncrona, mas poderá só ser resolvido algum tempo depois. Contudo, embora `Promise` seja um único valor (que pode ser um objeto de matriz), um `Observable` é uma matriz de objetos (possivelmente com apenas um único membro). Isso permite que o código chame [métodos de matriz](https://www.w3schools.com/jsref/jsref_obj_array.asp), como `concat`, `map` e `filter`, em objetos `Observable`.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p111">An `Observable` is like a `Promise` object in some ways - it is returned immediately from an asynchronous call, but it might not resolve until some time later. However, while a `Promise` is a single value (which can be an array object), an `Observable` is an array of objects (possibly with only a single member). This enables code to call [array methods](https://www.w3schools.com/jsref/jsref_obj_array.asp), such as `concat`, `map`, and `filter`, on `Observable` objects.</span></span> 

### <a name="pushing-instead-of-pulling"></a><span data-ttu-id="e53dc-149">Obter em vez de enviar</span><span class="sxs-lookup"><span data-stu-id="e53dc-149">Pushing instead of pulling</span></span>

<span data-ttu-id="e53dc-p112">Seu código "obtém" objetos `Promise` atribuindo-os a variáveis, mas objetos `Observable` "enviam" seus valores para objetos que se *inscrevem* no `Observable`. Os assinantes são objetos `Observer`. O benefício da arquitetura push é que novos membros podem ser adicionados à matriz `Observable` ao longo do tempo. Quando um novo membro é adicionado, todos os objetos `Observer` que assinam o `Observable` recebem uma notificação.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p112">Your code "pulls" `Promise` objects by assigning them to variables, but `Observable` objects "push" their values to objects that *subscribe* to the `Observable`. The subscribers are `Observer` objects. The benefit of the push architecture is that new members can be added to the `Observable` array over time. When a new member is added, all the `Observer` objects that subscribe to the `Observable` receive a notification.</span></span> 

<span data-ttu-id="e53dc-p113">O `Observer` é configurado para processar cada novo objeto (chamado o "próximo" objeto) com uma função. (Ele também é configurado para responder a um erro e a uma notificação de conclusão. Consulte a próxima seção para obter um exemplo.) Por esse motivo, os objetos `Observable` podem ser usados em uma maior variedade de cenários do que os objetos `Promise`. Por exemplo, além de retornarem um `Observable` de uma chamada AJAX, a maneira como você pode retornar um `Promise`, um `Observable` pode ser retornado de um manipulador de eventos, como o manipulador de eventos "modificado" de uma caixa de texto. Cada vez que um usuário insere texto na caixa, todos os objetos `Observer` inscritos reagem imediatamente usando o texto mais recente e/ou o estado atual do aplicativo como entrada.</span><span class="sxs-lookup"><span data-stu-id="e53dc-p113">The `Observer` is configured to process each new object (called the "next" object) with a function. (It is also configured to respond to an error and a completion notification. See the next section for an example.) For this reason, `Observable` objects can be used in a wider range of scenarios than `Promise` objects. For example, in addition to returning an `Observable` from an AJAX call, the way you can return a `Promise`, an `Observable` can be returned from an event handler, such as the "changed" event handler for a text box. Each time a user enters text in the box, all the subscribed `Observer` objects react immediately using the latest text and/or the current state of the application as input.</span></span> 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a><span data-ttu-id="e53dc-159">Aguardando a conclusão de todas as chamadas assíncronas</span><span class="sxs-lookup"><span data-stu-id="e53dc-159">Waiting until all asynchronous calls have completed</span></span>

<span data-ttu-id="e53dc-160">Quando quiser garantir que um retorno de chamada só seja executado quando todos os membros de um conjunto de objetos `Promise` forem resolvidos, use o método `Promise.all()`.</span><span class="sxs-lookup"><span data-stu-id="e53dc-160">When you want to ensure that a callback only runs when every member of a set of `Promise` objects has resolved, use the `Promise.all()` method.</span></span>

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
``` 

<span data-ttu-id="e53dc-161">Para fazer o mesmo com um objeto `Observable`, use o método [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).</span><span class="sxs-lookup"><span data-stu-id="e53dc-161">To do the same thing with an `Observable` object, you use the [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) method.</span></span>  

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

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a><span data-ttu-id="e53dc-162">Compilar o aplicativo Angular usando o compilador AOT (Ahead-of-Time)</span><span class="sxs-lookup"><span data-stu-id="e53dc-162">Compile the Angular application using the Ahead-of-Time (AOT) compiler</span></span>

<span data-ttu-id="e53dc-163">O desempenho do aplicativo é um dos aspectos mais importantes da experiência do usuário.</span><span class="sxs-lookup"><span data-stu-id="e53dc-163">Application performance is one of the most important aspects of user experience.</span></span> <span data-ttu-id="e53dc-164">Um aplicativo Angular pode ser otimizado usando o compilador Angular AOT (Ahead-of-Time) para compilar o aplicativo durante a compilação.</span><span class="sxs-lookup"><span data-stu-id="e53dc-164">An Angular application can be optimized by using the Angular Ahead-of-Time (AOT) compiler to compile the app at build time.</span></span> <span data-ttu-id="e53dc-165">Ele converte todo o código-fonte (modelos HTML e TypeScript) em um código JavaScript eficiente.</span><span class="sxs-lookup"><span data-stu-id="e53dc-165">It converts all source code (HTML templates and TypeScript) into efficient JavaScript code.</span></span> <span data-ttu-id="e53dc-166">Se você compilar o aplicativo com o compilador AOT, nenhuma compilação adicional ocorrerá no tempo de execução, o que resultará em um processamento mais rápido e solicitações assíncronas mais rápidas para modelos HTML.</span><span class="sxs-lookup"><span data-stu-id="e53dc-166">If you compile your app with the AOT compiler, no additional compilation will occur at runtime, which results in faster rendering and faster asynchronous requests for HTML templates.</span></span> <span data-ttu-id="e53dc-167">Além disso, o tamanho geral do aplicativo diminui, pois o compilador Angular não precisa ser incluído no aplicativo para distribuição.</span><span class="sxs-lookup"><span data-stu-id="e53dc-167">Additionally, the overall application size will be reduced, because the Angular compiler won't need to be included in the application distributable.</span></span> 

<span data-ttu-id="e53dc-168">Para usar o compilador AOT, adicione `--aot` aos comandos `ng build` ou `ng serve`:</span><span class="sxs-lookup"><span data-stu-id="e53dc-168">To use the AOT compiler, add `--aot` to the `ng build` or `ng serve` command:</span></span>

```bash
ng build --aot
ng serve --aot
```

> [!NOTE]
> <span data-ttu-id="e53dc-169">Para saber mais sobre o compilador Angular AOT (Ahead-of-Time), consulte o [guia oficial](https://angular.io/guide/aot-compiler).</span><span class="sxs-lookup"><span data-stu-id="e53dc-169">To learn more about the Angular Ahead-of-Time (AOT) compiler, see the [official guide](https://angular.io/guide/aot-compiler).</span></span>
