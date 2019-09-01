# <a name="contribute-to-this-documentation"></a>Contribuir para esta documentação

Agradecemos seu interesse em nossa documentação!

* [Maneiras de contribuir](#ways-to-contribute)
* [Contribuir usando o GitHub](#contribute-using-github)
* [Contribuir usando o Git](#contribute-using-git)
* [Como usar o Markdown para formatar seu tópico](#how-to-use-markdown-to-format-your-topic)
* [Perguntas frequentes](#faq)
* [Mais recursos](#more-resources)

## <a name="ways-to-contribute"></a>Maneiras de contribuir

Veja algumas maneiras de contribuir com esta documentação:

* Para fazer pequenas alterações em um artigo, [contribuir usando o GitHub](#contribute-using-github).
* Para fazer grandes alterações ou alterações que envolvem código, [contribuir usando o Git](#contribute-using-git).
* Relatar erros de documentação por meio de problemas do GitHub.
* Solicitar nova documentação no site [UserVoice da plataforma de desenvolvedor do Office](http://officespdev.uservoice.com) .

## <a name="contribute-using-github"></a>Contribuir usando o GitHub

Use o GitHub para contribuir para esta documentação sem precisar clonar o repositório na sua área de trabalho. Essa é a maneira mais fácil de criar uma solicitação pull neste repositório. Use este método para fazer uma alteração menor que não envolve alterações de código. 

**Observação**: o uso desse método permite contribuir com um artigo de cada vez.

### <a name="to-contribute-using-github"></a>Para contribuir usando o GitHub

1. Encontre o artigo que você deseja contribuir no GitHub.
2. Quando estiver no artigo no GitHub, entre no GitHub (obtenha uma conta gratuita do [GitHub](https://github.com/join)).
3. Escolha o **ícone de lápis** (edite o arquivo na bifurcação deste projeto) e faça as alterações na janela **<>Editar arquivo** . 
4. Role até a parte inferior e insira uma descrição.
5. Escolha **propor alteração**>de arquivo**criar solicitação pull**.

Agora você enviou com êxito uma solicitação pull. Solicitações pull geralmente são analisadas dentro de 10 dias úteis. 


## <a name="contribute-using-git"></a>Contribuir usando o Git

Use o Git para contribuir com alterações substantivas, como:

* Código contribuinte.
* Contribuindo alterações que afetam o significado.
* Contribuindo com grandes alterações em texto.
* Adição de novos tópicos.

### <a name="to-contribute-using-git"></a>Para contribuir usando o Git

1. Se você não tiver uma conta do GitHub, configure uma no [GitHub](https://github.com/join). 
2. Depois de ter uma conta, instale o Git em seu computador. Siga as etapas no tutorial de [configuração do git] .
3. Para enviar uma solicitação pull usando o Git, siga as etapas em [usar o GitHub, o git e este repositório](#use-github-git-and-this-repository).
4. Você será solicitado a assinar o contrato de licença do colaborador se for:

    * Um membro do grupo Microsoft Open Technologies.
    * Um colaborador que não funciona para a Microsoft.

Como membro da Comunidade, você deve assinar o contrato de licença de contribuição (CLA) antes de poder contribuir em grandes envios para um projeto. Você precisa concluir e enviar a documentação apenas uma vez. Reveja cuidadosamente o documento. Talvez seja necessário que seu empregador assine o documento.

A assinatura do CLA não concede direitos para confirmar o repositório principal, mas significa que as equipes de publicação de conteúdo do desenvolvedor do Office e do Office Developer poderão revisar e aprovar suas contribuições. Você é creditado nos seus envios.

Solicitações pull geralmente são analisadas dentro de 10 dias úteis.

## <a name="use-github-git-and-this-repository"></a>Use o GitHub, o Git e este repositório

**Observação**: a maior parte das informações desta seção pode ser encontrada nos artigos de [ajuda do GitHub] .  Se você estiver familiarizado com o git e o GitHub, pule para a seção **contribuir e editar conteúdo** para as especificações do fluxo de código/conteúdo desse repositório.

### <a name="to-set-up-your-fork-of-the-repository"></a>Para configurar sua bifurcação do repositório

1.  Configure uma conta GitHub para que você pode contribuir para esse projeto. Se você ainda não fez isso, vá para o [GitHub](https://github.com/join) e faça isso agora.
2.  Instale o Git em seu computador. Siga as etapas no tutorial de [configuração do git] .
3.  Crie sua própria bifurcação para este repositório. Para fazer isso, na parte superior da página, escolha o botão **** de bifurcação.
4.  Copie sua bifurcação para seu computador. Para fazer isso, abra o Git bash. No prompt de comando, digite:

        git clone https://github.com/<your user name>/<repo name>.git

    Em seguida, crie uma referência para o repositório raiz inserindo esses comandos:

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Parabéns! Agora seu repositório está configurado. Você não precisará repetir essas etapas novamente.

### <a name="contribute-and-edit-content"></a>Contribuir e editar o conteúdo

Para tornar o processo de contribuição o mais simples possível, siga estas etapas.

#### <a name="to-contribute-and-edit-content"></a>Para contribuir e editar conteúdo

1. Crie uma nova ramificação.
2. Adicione novo conteúdo ou edite o conteúdo existente.
3. Envie uma solicitação pull para o repositório principal.
4. Exclua a ramificação.

**Importante** Limite cada filial a um único conceito/artigo para simplificar o fluxo de trabalho e reduzir a chance de conflitos de mesclagem. O conteúdo apropriado para uma nova ramificação inclui:

* Um novo artigo.
* Edição ortográfica e gramatical.
* Aplicar uma única alteração de formatação em um grande conjunto de artigos (por exemplo, aplicando um novo rodapé de direitos autorais).

#### <a name="to-create-a-new-branch"></a>Para criar uma nova ramificação

1.  Abra o Git bash.
2.  No prompt de comando do git bash, `git pull upstream master:<new branch name>`digite. Isso cria uma nova ramificação localmente que é copiada da ramificação mestra OfficeDev mais recente.
3.  No prompt de comando do git bash, `git push origin <new branch name>`digite. Isso alerta o GitHub para a nova ramificação. Agora você deverá surgir a nova ramificação na sua bifurcação do repositório no GitHub.
4.  No prompt de comando do git bash, `git checkout <new branch name>` digite para mudar para a nova ramificação.

#### <a name="add-new-content-or-edit-existing-content"></a>Adicionar novo conteúdo ou editar o conteúdo existente

Navegue até o repositório no computador usando o explorador de arquivos. Os arquivos do repositório estão `C:\Users\<yourusername>\<repo name>`em.

Para editar arquivos, abra-os em um editor de sua escolha e modifique-os. Para criar um novo arquivo, use o editor de sua escolha e salve o novo arquivo no local apropriado na sua cópia local do repositório. Enquanto estiver trabalhando, salve seu trabalho com frequência.

Os arquivos no `C:\Users\<yourusername>\<repo name>` são uma cópia de trabalho da nova ramificação que você criou em seu repositório local. Qualquer que seja a alteração você faça nessa pasta, ela só afetará o repositório local quando você confirmar uma alteração. Para confirmar uma alteração no repositório local, digite os seguintes comandos no GitBash:

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

O comando `add` adiciona suas alterações para uma área de preparo em preparação para confirmá-las no repositório. O período após o `add` comando especifica que você deseja testar todos os arquivos adicionados ou modificados, verificando as subpastas recursivamente. (Caso você não queira confirmar todas as alterações, é possível adicionar arquivos específicos. Você também pode desfazer uma confirmação. Para obter ajuda, digite `git add -help` ou `git status`.)

O comando `commit` aplica as alterações preparadas ao repositório. A opção `-m` significa que você está fornecendo o comentário de confirmação na linha de comando. As opções-v e-a podem ser omitidas. A opção-v é para a saída detalhada do comando e-a faz o que você já fez com o comando adicionar.

Você pode confirmar várias vezes enquanto estiver fazendo seu trabalho ou pode confirmar uma vez quando terminar.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Enviar uma solicitação pull para o repositório principal.

Quando você tiver concluído o trabalho e estiver pronto para mesclá-lo no repositório principal, siga estas etapas.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Para enviar uma solicitação pull para o repositório principal

1.  No prompt de comando do git bash, `git push origin <new branch name>`digite. Em seu repositório local, `origin` refere-se ao repositório do GitHub a partir do qual você clonou o repositório local. Esse comando coloca o estado atual do sua nova ramificação, incluindo todas as confirmações feitas nas etapas anteriores, na ramificação do GitHub.
2.  No site do GitHub, navegue em sua bifurcação para a nova ramificação.
3.  Escolha o botão de **solicitação pull** na parte superior da página.
4.  Verifique se o Branch base `OfficeDev/<repo name>@master` é e a ramificação `<your username>/<repo name>@<branch name>`principal é.
5.  Escolha o botão **Atualizar intervalo de confirmação** .
6.  Adicione um título à sua solicitação pull e descreva todas as alterações que você está fazendo.
7.  Envie a solicitação pull.

Um dos administradores do site processará sua solicitação pull. Sua solicitação pull será a área de OfficeDev/<repo name> site em problemas. Quando a solicitação pull for aceita, o problema será resolvido.

#### <a name="create-a-new-branch-after-merge"></a>Criar uma nova ramificação após a mesclagem

Depois que uma ramificação é mesclada com êxito (ou seja, sua solicitação pull é aceita), não continue trabalhando nessa filial local. Isso pode levar a conflitos de mesclagem se você enviar outra solicitação pull. Para fazer outra atualização, crie uma nova ramificação local da ramificação mesclada com êxito e exclua sua ramificação local inicial.

Por exemplo, se a sua filial local X foi mesclada com êxito no Branch mestre OfficeDev/Microsoft-Graph-docs e você deseja fazer atualizações adicionais para o conteúdo que foi mesclado. Crie uma nova ramificação local, X2, da ramificação mestre OfficeDev/Microsoft-Graph-docs. Para fazer isso, abra GitBash e execute os seguintes comandos:

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Agora você tem cópias locais (em uma nova filial local) do trabalho que você enviou no Branch X. A ramificação X2 também contém todos os outros escritores de trabalho mesclados, portanto, se o trabalho depender de trabalho de outras pessoas (por exemplo, imagens compartilhadas), ele estará disponível na nova ramificação. Você pode verificar se o trabalho anterior (e outras pessoas) está na ramificação fazendo check-out da nova ramificação...

    git checkout X2

... e verificando o conteúdo. (O `checkout` comando atualiza os arquivos no `C:\Users\<yourusername>\microsoft-graph-docs` estado atual da ramificação X2.) Depois de fazer o check-out da nova ramificação, você pode fazer atualizações no conteúdo e confirmá-las como de costume. No entanto, para evitar trabalhar na ramificação mesclada (X) por engano, o melhor a fazer será excluí-la (confira a seguinte seção: **Excluir uma ramificação**).

#### <a name="delete-a-branch"></a>Excluir uma ramificação

Depois que as alterações forem mescladas com êxito no repositório principal, exclua a ramificação usada porque você não precisa mais dela.  Qualquer trabalho adicional deve ser feito em uma nova ramificação.  

#### <a name="to-delete-a-branch"></a>Para excluir uma ramificação

1.  No prompt de comando do git bash, `git checkout master`digite. Isso garante que você não fique na ramificação a ser excluída (o que não é permitido).
2.  Em seguida, no prompt de comando, `git branch -d <branch name>`digite. Isso excluirá a ramificação em seu computador somente se ela tiver sido mesclada com êxito no repositório upstream. (Você pode substituir esse comportamento com o sinalizador `–D`, mas primeiro certifique-se de que deseja fazer isso).
3.  Por fim, digite `git push origin :<branch name>` no comando prompt (um espaço antes dos dois pontos e nenhum espaço depois deles).   Essa ação excluirá a ramificação em uma bifurcação do github.  

Parabéns, você contribuiu com êxito para o projeto!

## <a name="how-to-use-markdown-to-format-your-topic"></a>Como usar o Markdown para formatar seu tópico

### <a name="markdown"></a>Markdown

Todos os artigos neste repositório usam Markdown. Uma introdução completa (e a listagem de todas as sintaxes) podem ser encontradas em [Daring Fireball – suredução].
 
## <a name="faq"></a>Perguntas frequentes

### <a name="how-do-i-get-a-github-account"></a>Como obter uma conta do GitHub?

Preencha o formulário em [Ingressar no GitHub](https://github.com/join) para abrir uma conta gratuita do GitHub. 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Onde posso obter um Contrato de Licença do Colaborador? 

Um aviso será automaticamente enviado para você informando que é preciso assinar o CLA (Contrato de Licença do Colaborador) se sua solicitação de recebimento exigir um. 

Como membro da comunidade, **você deve assinar o CLA (Contrato de Licença do Colaborador) antes de poder contribuir com envios volumosos para esse projeto**. Você só precisa concluir e enviar a documentação uma vez. Reveja cuidadosamente o documento. Talvez seja necessário que seu empregador assine o documento.

### <a name="what-happens-with-my-contributions"></a>O que acontece com minhas contribuições?

Quando você enviar suas alterações, por meio de uma solicitação pull, nossa equipe será notificada e revisará sua solicitação pull. Você receberá notificações sobre sua solicitação pull do GitHub; Você também pode ser notificado por alguém da nossa equipe se precisar de mais informações. Se sua solicitação pull for aprovada, atualizaremos a documentação. Reservamos o direito de editar seu envio por questões jurídicas, de estilo, de clareza ou de outros problemas.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Posso se tornar um Aprovador para as solicitações pull do GitHub do repositório?

No momento, não estamos permitindo que colaboradores externos aprovem solicitações pull neste repositório.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Em quanto tempo receberei uma resposta sobre minha solicitação de alteração?

Solicitações pull geralmente são analisadas dentro de 10 dias úteis.


## <a name="more-resources"></a>Mais recursos

* Para saber mais sobre redução, acesse o site do criador de redução [Daring Fireball].
* Para saber mais sobre como usar o git e o GitHub, primeiro Confira a [ajuda do GitHub].

[GitHub Home]: http://github.com
[Ajuda do GitHub]: http://help.github.com/
[Configurar o Git]: https://help.github.com/articles/set-up-git/
[Daring Fireball]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
