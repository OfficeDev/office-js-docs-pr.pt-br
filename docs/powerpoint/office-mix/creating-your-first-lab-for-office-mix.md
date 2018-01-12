
# <a name="walkthrough-creating-your-first-lab-for-office-mix"></a>Passo a passo: criar o seu primeiro laboratório para o Office Mix
Crie seu primeiro laboratório LabsJS usando instruções passo a passo.



Neste passo a passo você criará um laboratório LabsJS simples do zero. Seu laboratório será um questionário simples com as opções de verdadeiro/falso e que fornece apenas uma única pergunta. 

Em vez de começar com um modelo de projeto do Visual Studio, você começará com três arquivos vazios; isso mostra o quão simples um laboratório é: 


- TrueFalse.html (html5)
    
- TrueFalse.js
    
- TrueFalse.css
    
Você pode usar qualquer editor de código que quiser para editar esses arquivos, pois não começaremos com um modelo do Visual Studio. Na verdade, o arquivo HTML é simples e, se você quiser, poderá copiar/colar a marcação HTML dos arquivos do tutorial. Entretanto, observe que ele deve ser HTML5. Portanto, certifique-se de que sua declaração de tipo de documento seja `<!DOCTYPE html>`. O arquivo CSS é opcional. Todo o trabalho pesado é feito no arquivo JavaScript (.js), TrueFalse.js. O passo a passo cobrirá quatro recursos principais de laboratório:

- Configuração (conexão com o host)
    
- Alterações de modo (entre o modo de edição e o modo de exibição)
    
- Edição do laboratório
    
- Obtenção (ou execução) do laboratório
    

 **Observação**  
 ---
 O arquivo Labhost.html é executado em um servidor Web e fornece o ambiente de hospedagem para o desenvolvimento e teste do laboratório. Isso simplifica bastante o desenvolvimento do laboratório. Confira [Introdução ao LabsJS para Office Mix](get-started-with-labsjs-for-office-mix.md) para saber mais sobre como configurar seu ambiente de desenvolvimento.<br/><br/>

Finalmente, você pode exibir o arquivo JavaScript completo (TrueFalse.js) entre os arquivos distribuídos com este SDK. O que vem a seguir é um passo a passo do processo de codificação.

## <a name="connecting-to-the-lab-host"></a>Conexão com o host do laboratório

Os laboratórios nesse ambiente são capazes de executar com nosso host de laboratório (para desenvolvimento e teste) ou com o host de tempo de execução padrão fornecido pelo host Office.js. A função de abertura usa uma expressão simples if/else para testar quais desses contextos de hospedagem se aplicam.


```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```

O objeto **PostMessageLabHost** é executado no ambiente de desenvolvimento labhost.html, enquanto na produção o laboratório é executado no PowerPoint/Office Mix usando o **OfficeJSLabHost**.

Em seguida, crie um método auxiliar para gerar um retorno de chamada cujo trabalho é resolver ou rejeitar um objeto jQuery adiado que você passou. Use o método **createCallback** para ir do jQuery local até os retornos de chamada definidos pelo labs.js.




```js
function createCallback(deferred) {
    return function (err, data) {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```

Também criamos um método auxiliar para recuperar a configuração do laboratório para uma determinada pergunta e resposta.




```js
function getConfiguration(question, answer) {
    var choiceComponent = {
        name: question,
        type: Labs.Components.ChoiceComponentType,
        timeLimit: 0,
        maxAttempts: 1,
        choices: [
            { id: "0", name: "True", value: "True" },
            { id: "1", name: "False", value: "False" }],
        maxScore: 1,
        hasAnswer: true,
        answer: answer ? "0" : "1",
        values: null,
        secure: false,
        data: null
    };

    return {
        appVersion: { major: 0, minor: 1 },
        components: [choiceComponent],
        name: question,
        timeline: null,
        analytics: null
    };
}
```


## <a name="mode-changes"></a>Alterações de modo

Um laboratório está sempre em um dentro dois estados ou modos:  **exibição** e **edição**. Portanto, precisamos de uma maneira de capturar e manter o estado e o comportamento para o teste; vamos criar uma classe para essa finalidade.


```js
var TrueFalseQuiz = (function () {
    /**
     * Constructor - takes in the starting mode.
     */
    function TrueFalseQuiz(mode) {
        var self = this;        
        self._modeSwitchP = $.when();
        self._labInstance = null;
        self._labEditor = null;        
      /**
       * Listen for mode changed events and 
       * then switch accordingly. Also set the initial mode state.
       */
        Labs.on(Labs.Core.EventTypes.ModeChanged, function (modeChangedEvent) {
            self.switchUserMode(Labs.Core.LabMode[modeChangedEvent.mode]);
        });
        this.switchUserMode(mode);        
    }
```

Além disso, fornecemos um método auxiliar cujo trabalho é atualizar a interface do usuário do teste com base em se a resposta de uma pergunta do teste (ou seja, o "envio") está correta ou incorreta.




```js
    TrueFalseQuiz.prototype._showResults = function(correct) {
        $("#submit-button").removeClass("btn-default");
        $("#submit-button").addClass(correct ? "btn-success" : "btn-danger");
        $("#submit-button").text(correct ? "Correct!" : "Incorrect");

        $("#submit-button").prop("disabled", true);
        $("input:radio[name='quizAnswers']").prop("disabled", true);
    };
```

Também precisamos de uma função para alternar entre modos de exibição e de edição.




```js
TrueFalseQuiz.prototype.switchUserMode = function (mode) {
        var self = this;

        // Wait for any previous mode switch to complete before performing the new one
        self._modeSwitchP = self._modeSwitchP.then(function () {
            var switchedStateDeferred = $.Deferred();

            // Clean up any variables associated with the previous mode.
            if (self._labInstance) {
                $("#quiz-view-form").off("submit");
                self._labInstance.done(createCallback(switchedStateDeferred));
            } else if (self._labEditor) {
                self._unbindFromEditUpdates();
                self._labEditor.done(createCallback(switchedStateDeferred));
            } else {
                switchedStateDeferred.resolve();
            }

            // After the cleanup occurs, switch to the new mode.
            return switchedStateDeferred.promise().then(function () {
                self._labEditor = null;
                self._labInstance = null;

                if (mode === Labs.Core.LabMode.Edit) {
                    return self._switchToEditMode();
                } else {
                    return self._switchToViewMode();
                }
            });
        });

        // Display an error if it occurs.
        self._modeSwitchP.fail(function (error) {
            // ... error handling ...
        });
    };
```

Nossa próxima função atualiza a configuração do teste com base em eventos de alteração que recebemos na interface do usuário.




```js
    TrueFalseQuiz.prototype._updateConfigurationFromUI = function () {
        var question = $("#question-edit").val();
        var answerIsTrue = $("input:radio[name='answerValue']:checked").val() === "true";

        this._updateConfiguration(question, answerIsTrue, true, function (err) {
            if (err) {
                // show error
            }
        });
    };
```

Depois, atualizamos os dados de configuração do laboratório armazenados no servidor com base em determinadas perguntas e respostas.




```js
    TrueFalseQuiz.prototype._updateConfiguration = function (question, answer, serialize, callback) {
        var configuration = getConfiguration(question, answer);

        if (serialize) {
            this._labEditor.setConfiguration(configuration, callback);
        } else {
            callback(null, null);
        }
    };
```

Em seguida, temos uma função que vincula as atualizações feitas no laboratório no modo de edição às alterações de configuração que fizemos. Depois disso vem o código para desvinculação dos manipuladores de alteração associados anteriormente.




```js
    TrueFalseQuiz.prototype._bindToEditUpdates = function () {
        var self = this;

        // Listen for the question changing
        $("#question-edit").on("input propertychange paste", function () {
            self._updateConfigurationFromUI();
        });

        $('input[name="answerValue"]').on("change", function (e) {
            self._updateConfigurationFromUI();
        });
    };
```




```js
    TrueFalseQuiz.prototype._unbindFromEditUpdates = function () {
        $("#question-edit").off("input propertychange paste");
        $('input[name="answerValue"]').off("change");
    };
```

Agora vem uma parte fundamental da seção, ou seja, os métodos para alternar entre os modos de exibição e de edição. Vamos começar alternando do modo de exibição para o modo de edição.




```js
    TrueFalseQuiz.prototype._switchToEditMode = function () {
        var self = this;
        var editLabDeferred = $.Deferred();

        // Make the Labs.js API call to edit the lab.
        Labs.editLab(createCallback(editLabDeferred));

        return editLabDeferred.promise().then(function (labEditor) {            
            self._labEditor = labEditor;

            // Retrieve any existing configuration from the lab editor.
            var configurationDeferred = $.Deferred();
            labEditor.getConfiguration(createCallback(configurationDeferred));

            return configurationDeferred.promise().then(function (configuration) {
                var configurationReadyDeferred = $.Deferred();

                // Get the question and answer values if they exist. 
                //Otherwise use the defaults.
                var question = configuration !== null ? configuration.components[0].name : "";
                var answerIsTrue = configuration !== null ? configuration.components[0].answer === "0" : true;

                // Update the lab configuration based on the question and answer.
                self._updateConfiguration(
                    question,
                    answerIsTrue,
                    configuration === null,
                    createCallback(configurationReadyDeferred));

                // Update the UI based on the question and answer.
                $("#question-edit").val(question);
                $('input[name="answerValue"][value="' + answerIsTrue + '"]').prop('checked', true);

                // Bind to changes.
                self._bindToEditUpdates();

                // Flip over the UI.
                $("#quiz-editor").removeClass("hidden");
                $("#quiz-view").addClass("hidden");

                return configurationReadyDeferred.promise();
            });
        });
    };
```

E agora, alternando do modo de edição para modo de exibição.




```js
    TrueFalseQuiz.prototype._switchToViewMode = function () {
        var self = this;
        var takeLabDeferred = $.Deferred();

        // Call the labs.js API to start taking the lab.
        Labs.takeLab(createCallback(takeLabDeferred));

        return takeLabDeferred.promise().then(function (labInstance) {
            self._labInstance = labInstance;

            // Get the choice component instance that will be generated
            // from the choice component we saved when editing the lab.
            var choiceComponentInstance = self._labInstance.components[0];

            // Get the attempts associated with that choice component.
            var attemptsDeferred = $.Deferred();
            choiceComponentInstance.getAttempts(createCallback(attemptsDeferred));
            var attemptP = attemptsDeferred.promise().then(function (attempts) {
                // See if we already had started an attempt against 
                // the problem. If not create one.
                var currentAttemptDeferred = $.Deferred();
                if (attempts.length > 0) {
                    currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
                } else {
                    choiceComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
                }

                return currentAttemptDeferred.then(function (currentAttempt) {
                    var resumeDeferred = $.Deferred();

                    // After we have the attempt, mark that we are resuming
                    // it as well. This will note the resumption time
                    // in the lab activity log.
                    currentAttempt.resume(createCallback(resumeDeferred));
                    return resumeDeferred.promise().then(function () {
                        return currentAttempt;
                    });
                });
            });

            return attemptP.promise().then(function (attempt) {
                // Store off the latest attempt for later use.
                self._currentAttempt = attempt;

                // Update the question field of the view UI.
                $("#question-view").text(choiceComponentInstance.component.name);

                // Determine whether the quiz has already been taken
                // and update the UI accordingly.
                var submissions = attempt.getSubmissions();
                if (submissions.length > 0) {
                    var correctAttempt = submissions[submissions.length - 1].result.score === 1;
                    var submissionValue = submissions[submissions.length - 1].answer.answer === "0";
                    $('input[name="quizAnswers"][value="' + submissionValue + '"]').prop('checked', true);
                    self._showResults(correctAttempt);
                } else {
                    $("#submit-button").removeClass("btn-success btn-danger"    );
                    $("#submit-button").addClass("btn-default");
                    $("#submit-button").text("Submit");
                    $("#submit-button").prop("disabled", false);
                    $("input:radio[name='quizAnswers']").prop("disabled", false);
                }                

                // Hook up the form submit button and then
                // grade the attempt when it is selected.
                $("#quiz-view-form").on("submit", function (e) {
                    e.preventDefault();
                    
                    // Get the checked value and see whether the choice
                    // was true or false - map back to our choice fields.
                    var submission = $("input:radio[name='quizAnswers']:checked").val() === "true" ? "0" : "1";

                    // Grade against the stored answer.
                    var correct = choiceComponentInstance.component.answer === submission;

                    // Submit the attempt with the labs.js API.
                    attempt.submit(
                        new Labs.Components.ChoiceComponentAnswer(submission),
                        new Labs.Components.ChoiceComponentResult(correct ? 1 : 0, true),
                        function (err) {
                            if (err) {
                                // Error
                            }
                        });

                    // And finally update the UI.
                    self._showResults(correct);
                });

                // And make the view UI visible.
                $("#quiz-editor").addClass("hidden");
                $("#quiz-view").removeClass("hidden");
            });
        });
    };

    return TrueFalseQuiz;
})();
```

Por fim, após a conexão com o host e finalização do documento, inicie o teste.




```js
$(document).ready(function () {
    Labs.connect(function (err, connectionResponse) {
        if (err) {
            // ... error handling goes here ...
            return;
        }

        // Start up the true/false quiz.
        var trueFalseQuiz = new TrueFalseQuiz(connectionResponse.mode);
    });
});
```


## <a name="additional-resources"></a>Recursos adicionais
<a name="bk_addresources"> </a>


- [Suplementos do Office Mix](office-mix-add-ins.md)
    
