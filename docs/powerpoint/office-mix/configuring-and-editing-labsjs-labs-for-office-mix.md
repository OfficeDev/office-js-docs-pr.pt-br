
# <a name="configuring-and-editing-labsjs-labs-for-office-mix"></a>Configuração e edição de laboratórios LabsJS para o Office Mix



O Office Mix fornece métodos do office.js para obter e definir configurações de laboratório. A configuração indica ao Office Mix que tipo de laboratório você está criando, bem como o tipo de dados que o laboratório enviará de volta. Essas informações são usadas para coletar e visualizar análises.

## <a name="getting-the-lab-editor"></a>Como obter o editor de laboratório

O editor de laboratório, o objeto [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md), permite que você edite seu laboratório e obtenha e defina a configuração de seu laboratório. Ao terminar de editar seu laboratório, você precisa chamar o método **Done**. No entanto, chamar o método **Done** não é necessário, exceto quando você está tentando obter ou executar um laboratório que você está editando. Observe que é possível abrir apenas uma instância do laboratório por vez.

O código a seguir mostra como obter o editor de laboratório.




```js
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

Use os métodos **getConfiguration** e **setConfiguration** no [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md) para armazenar a configuração de um determinado laboratório. A configuração ([Labs.Core.IConfiguration](../../../reference/office-mix/labs.core.iconfiguration.md)) indica ao Office Mix quais dados serão coletados e processados pelo laboratório. Uma configuração contém informações gerais sobre um laboratório, incluindo nome, versão e outras opções de configuração. A parte mais importante da configuração é a definição dos componentes de laboratório.

O código a seguir mostra como definir e obter uma configuração. Para definir uma configuração, basta criar o objeto da configuração e, em seguida, chamar o método **setConfiguration**. Para recuperar depois a configuração, chame o método **getConfiguration** no objeto do editor de laboratório.




```js

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```js

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## <a name="closing-the-editor"></a>Como fechar o editor

Para fechar o editor, chame o método **Done** no editor quando terminar de editar o laboratório. Observe que não é possível obter e editar um laboratório. Depois de chamar **Done**, você pode editar ou executar o laboratório.


## <a name="interacting-with-a-lab"></a>Como interagir com um laboratório

Depois de definir a configuração do laboratório, você está pronto para começar a interação com o laboratório. Quando o laboratório é executado dentro do PowerPoint, as interações são simuladas. No entanto, quando o laboratório é executado dentro do reprodutor de lição do Office Mix, os dados são armazenados no banco de dados do Office Mix e usados na análise.


### <a name="getting-the-lab-instance"></a>Como obter a instância do laboratório

Interaja com o laboratório usando o objeto [Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md), que é uma instância do laboratório configurado para o usuário atual. Para executar (ou "obter") o laboratório, chame a função [Labs.takeLab](../../../reference/office-mix/labs.takelab.md).


```js
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

O objeto da instância contém uma matriz de instâncias do componente ([Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md), [Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)) que mapeia para os componentes especificados na configuração. Na verdade, uma instância é simplesmente uma versão transformada da configuração usada para conectar IDs do lado do servidor a objetos da instância, bem como ocultar determinados campos do usuário quando for aplicável (por exemplo, dicas, respostas etc).


### <a name="managing-state"></a>Como gerenciar o estado

O estado é um armazenamento temporário associado a um usuário que está executando um determinado laboratório. Você pode usar o armazenamento para manter as informações entre chamadas sucessivas do laboratório. Por exemplo, um laboratório de programação pode armazenar o trabalho atual em andamento do usuário.

Para **definir** o estado, use o código a seguir.




```js
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

Para **obter** o estado, use o código a seguir.




```js
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## <a name="component-instances-and-results"></a>Instâncias e resultados do componente

Veja a seguir uma visão geral de como implementar as instâncias dos quatro tipos de componentes, e exemplos breves dos métodos do componente. 

Primeiro, no entanto, você precisa se familiarizar com dois conceitos importantes ao trabalhar com instâncias de componentes. O primeiro é o conceito de **tentativas** e **valores**.

 **Tentativas**

É uma tentativa por parte de um usuário para concluir uma instância do componente. Por exemplo, no caso de uma questão de múltipla escola, uma tentativa começa quando o usuário começa a solucionar o problema e termina quando uma pontuação final é atribuída. Em seguida, a análise do Office Mix coleta os resultados do usuário para o problema.


 >**Observação**:  As tentativas podem ser usadas para todos os tipos de componente, exceto para o tipo **DynamicComponent**.

Você pode recuperar os resultados de todas as tentativas associadas a uma certa instância de componente usando o método **getAttempts**. Depois de recuperar os resultados, o usuário pode tentar novamente uma das tentativas existentes usando o método **resume**, ou criar uma nova tentativa usando o método **createAttempt**. O exemplo a seguir mostra o processo.




```js
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **Values**

As instâncias do componente contêm um dicionário de chaves mapeado para uma matriz de valores. Você pode usar a matriz para armazenar dicas, comentários ou qualquer outro conjunto de valores que queira associar ao componente. A instância do componente fornece acesso a esses valores usando o método **getValues**.

A consulta a um valor de dica, por exemplo, faz com que a análise marque que o usuário usou uma dica. Os valores são controlados de acordo com a tentativa.

O exemplo de código a seguir mostra como consultar uma dica.




```js
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### <a name="activitycomponentinstance"></a>ActivityComponentInstance


Use o objeto **ActivityComponentInstace** para controlar a interação de um usuário com um componente de atividade. Essa classe fornece um método **complete** para indicar que o usuário terminou de interagir com a atividade. O método pode indicar que o usuário concluiu uma tarefa atribuída, concluiu uma leitura ou qualquer outro ponto de extremidade associado à atividade. O código a seguir mostra como usar o método **complete**.


```js
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### <a name="choicecomponentinstance"></a>ChoiceComponentInstance


Use o objeto **ChoiceComponentInstance** para controlar a interação do usuário com um componente de escolha. Os componentes de escolha são problemas que apresentam ao usuário uma lista de opções para seleção. Pode haver ou não uma resposta correta. A classe fornece dois métodos principais: **getSubmissions** e **submit**. O método **getSubmissions** permite que você recupere os envios armazenados anteriormente; o método **submit** permite que um novo envio seja armazenado. Os exemplos de código a seguir ilustram o uso dos métodos.


```js
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```js
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="inputcomponentinstance"></a>InputComponentInstance


Use o objeto **InputComponentInstance** para controlar a interação do usuário com um componente de entrada. A classe fornece dois métodos principais: **getSubmission** e **submit**. O método **getSubmissions** permite que você recupere os envios armazenados anteriormente; o método **submit** permite que você armazene um novo envio. O trecho de código a seguir ilustra o uso do método **getSubmissions**.


```js
var submissions = this._attempt.getSubmissions();
```

Ao usar o método **submit**, observe que o objeto **InputComponentAnswer** representa a resposta enviada, e o objeto **InputComponentResult** contém o resultado. O valor de retorno é um objeto **InputComponentSubmission** que contém a resposta, o resultado e um carimbo de hora que indica quando o resultado foi enviado.




```js
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="dynamiccomponentinstance"></a>DynamicComponentInstance


Use o objeto **DynamicComponentInstance** para controlar a interação do usuário com um componente dinâmico. Os métodos principais nesta classe são **getComponents**, **createComponent** e **close**.

O método **getComponents** permite que você recupere uma lista de instâncias de componente criadas anteriormente, conforme mostra o exemplo a seguir.




```js
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

O método **createComponent** constrói um novo componente e retorna essa instância do componente, conforme mostra o exemplo a seguir.




```js
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

Use o método **close** para indicar que você concluiu o uso do componente dinâmico para criar novos componentes. Observe que você também pode usar um método booliano **isClosed** para testar se a instância do componente dinâmico foi fechada. O código a seguir mostra como usar o método **close**.




```js
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## <a name="additional-resources"></a>Recursos adicionais



- [Suplementos do Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Passo a passo: Criando o primeiro laboratório para o Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
