# <a name="contribute-to-this-documentation"></a>Contribuir para esta documentação

Agradecemos seu interesse em nossa documentação!

* [Maneiras de contribuir](#ways-to-contribute)
* [Contribuir usando o GitHub](#contribute-using-github)
* [Contribuir usando o Git](#contribute-using-git)
* [Como usar o Markdown para formatar seu tópico](#how-to-use-markdown-to-format-your-topic)
* [Perguntas frequentes](#faq)
* [Mais recursos](#more-resources)

## <a name="ways-to-contribute"></a>Maneiras de contribuir

Aqui estão algumas maneiras de contribuir com esta documentação:

* Para fazer pequenas alterações em um artigo, [contribua usando GitHub](#contribute-using-github).
* Para fazer grandes alterações ou alterações que envolvam código, [Contribua usando Git](#contribute-using-git).
* Relatar bugs de documentação indo para a seção **Comentários** na parte inferior do artigo afetado e selecionando **Esta página** para criar um GitHub problema. Se isso não estiver disponível, crie um novo problema diretamente [no](https://github.com/OfficeDev/office-js-docs-pr/issues)GitHub .
* Solicitar nova documentação com [GitHub Problemas](https://github.com/OfficeDev/office-js-docs-pr/issues).

## <a name="contribute-using-github"></a>Contribuir usando o GitHub

Use GitHub para contribuir com essa documentação sem precisar clonar o repo para sua área de trabalho. Essa é a maneira mais fácil de criar uma solicitação pull neste repositório. Use este método para fazer uma pequena alteração que não envolva alterações de código.

**Observação:** o uso desse método permite que você contribua para um artigo de cada vez.

### <a name="to-contribute-using-github"></a>Para contribuir usando GitHub

1. Encontre o artigo para o GitHub.
2. Depois de entrar no artigo no GitHub, entre no GitHub (obter uma conta [gratuita Join GitHub](https://github.com/join)).
3. Escolha o **ícone de lápis** (edite o arquivo em sua bifurcação deste projeto) e faça as alterações na janela editar<>**de arquivo.**
4. Role até a parte inferior e insira uma descrição.
5. Escolha **Propor alteração de arquivo** Criar > **solicitação pull**.

Agora você enviou com êxito uma solicitação de pull. Normalmente, as solicitações de pull são revisadas dentro de 10 dias úteis.


## <a name="contribute-using-git"></a>Contribuir usando o Git

Use o Git para contribuir com alterações substantivas, como:

* Código de contribuição.
* Contribuir com alterações que afetam o significado.
* Contribuindo com grandes alterações no texto.
* Adicionando novos tópicos.

### <a name="to-contribute-using-git"></a>Para contribuir usando o Git

1. Se você não tiver uma conta GitHub, desarmá-la [em](https://github.com/join)GitHub .
2. Depois de ter uma conta, instale o Git no computador. Siga as etapas no tutorial [Configurar Git.]
3. Para enviar uma solicitação pull usando o Git, siga as etapas em [Usar GitHub, Git e este repositório.](#use-github-git-and-this-repository)
4. Você será solicitado a assinar o Contrato de Licença do Colaborador se estiver:

    * Um membro do grupo Microsoft Open Technologies.
    * Um colaborador que não trabalha para a Microsoft.

Como membro da comunidade, você deve assinar o Contrato de Licença de Contribuição (CLA) antes de poder contribuir com grandes envios para um projeto. Você só precisa concluir e enviar a documentação uma vez. Reveja cuidadosamente o documento. Talvez seja necessário que seu empregador assine o documento.

A assinatura do CLA não concede direitos de confirmação ao repositório principal, mas significa que as equipes de Publicação de Conteúdo de Desenvolvedor Office e Office desenvolvedores poderão revisar e aprovar suas contribuições. Você é creditado por seus envios.

Normalmente, as solicitações de pull são revisadas dentro de 10 dias úteis.

## <a name="use-github-git-and-this-repository"></a>Use o GitHub, o Git e este repositório

**Observação**: a maioria das informações nesta seção pode ser encontrada em GitHub [Artigos de] Ajuda.  Se você estiver familiarizado com Git e GitHub, pule para a seção Contribuir e **editar** conteúdo para as especificações do fluxo de código/conteúdo deste repositório.

### <a name="to-set-up-your-fork-of-the-repository"></a>Para configurar seu bifurcação do repositório

1. Configure uma conta GitHub para que você pode contribuir para esse projeto. Se você ainda não fez isso, vá [para](https://github.com/join) GitHub e faça isso agora.
2. Instale o Git em seu computador. Siga as etapas no tutorial [Configurar Git.]
3. Crie sua própria bifurcação para este repositório. Para fazer isso, na parte superior da página, escolha o **botão Bifurcação.**
4. Copie sua bifurcação para o computador. Para fazer isso, abra Git Bash. No prompt de comando, digite:

        git clone https://github.com/<your user name>/<repo name>.git

    Em seguida, crie uma referência para o repositório raiz inserindo esses comandos:

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Parabéns! Agora seu repositório está configurado. Você não precisará repetir essas etapas novamente.

### <a name="contribute-and-edit-content"></a>Contribuir e editar o conteúdo

Para tornar o processo de contribuição o mais contínuo possível, siga estas etapas.

#### <a name="to-contribute-and-edit-content"></a>Para contribuir e editar conteúdo

1. Crie uma nova ramificação.
2. Adicione novo conteúdo ou edite o conteúdo existente.
3. Envie uma solicitação pull para o repositório principal.
4. Exclua a ramificação.

**Importante** Limite cada filial a um único conceito/artigo para simplificar o fluxo de trabalho e reduzir a chance de conflitos de mesclagem. O conteúdo apropriado para um novo branch inclui:

* Um novo artigo.
* Edições ortográficas e gramaticais.
* Aplicar uma única alteração de formatação em um grande conjunto de artigos (por exemplo, aplicar um novo rodapé de direitos autorais).

#### <a name="to-create-a-new-branch"></a>Para criar um novo branch

1. Abra Git Bash.
2. No prompt de comando Git Bash, digite `git pull upstream master:<new branch name>` . Isso cria uma nova ramificação localmente copiada da filial mestra mais recente do OfficeDev.
3. No prompt de comando Git Bash, digite `git push origin <new branch name>` . Isso alerta GitHub para a nova filial. Agora você deverá surgir a nova ramificação na sua bifurcação do repositório no GitHub.
4. No prompt de comando Git Bash, digite `git checkout <new branch name>` para alternar para sua nova filial.

#### <a name="add-new-content-or-edit-existing-content"></a>Adicionar novo conteúdo ou editar o conteúdo existente

Navegue até o repositório em seu computador usando o Explorador de Arquivos. Os arquivos do repositório estão em `C:\Users\<yourusername>\<repo name>` .

Para editar arquivos, abra-os em um editor de sua escolha e modifique-os. Para criar um novo arquivo, use o editor de sua escolha e salve o novo arquivo no local apropriado em sua cópia local do repositório. Durante o trabalho, salve seu trabalho com frequência.

Os arquivos em `C:\Users\<yourusername>\<repo name>` são uma cópia de trabalho da nova ramificação que você criou no repositório local. Qualquer que seja a alteração você faça nessa pasta, ela só afetará o repositório local quando você confirmar uma alteração. Para confirmação de uma alteração no repositório local, digite os seguintes comandos no GitBash.

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

O comando `add` adiciona suas alterações para uma área de preparo em preparação para confirmá-las no repositório. O período após o comando especifica que você deseja estágio de todos os arquivos que você adicionou ou modificou, verificando `add` subpastas recursivamente. (Caso você não queira confirmar todas as alterações, é possível adicionar arquivos específicos. Você também pode desfazer uma confirmação. Para obter ajuda, digite `git add -help` ou `git status`.)

O comando `commit` aplica as alterações preparadas ao repositório. A opção `-m` significa que você está fornecendo o comentário de confirmação na linha de comando. As opções -v e -a podem ser omitidas. A opção -v é para saída detalhada do comando e -a faz o que você já fez com o comando add.

Você pode se comprometer várias vezes enquanto estiver fazendo seu trabalho ou pode se comprometer uma vez quando terminar.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Enviar uma solicitação pull para o repositório principal.

Quando terminar o trabalho e estiver pronto para mesclar no repositório principal, siga estas etapas.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Para enviar uma solicitação pull ao repositório principal

1. No prompt de comando Git Bash, digite `git push origin <new branch name>` . Em seu repositório local, `origin` refere-se ao repositório do GitHub a partir do qual você clonou o repositório local. Esse comando coloca o estado atual do sua nova ramificação, incluindo todas as confirmações feitas nas etapas anteriores, na ramificação do GitHub.
2. No site do GitHub, navegue em sua bifurcação para a nova ramificação.
3. Escolha o **botão Puxar Solicitação** na parte superior da página.
4. Verifique se o branch Base é `OfficeDev/<repo name>@master` e o branch Head é `<your username>/<repo name>@<branch name>` .
5. Escolha o **botão Atualizar Intervalo de Confirmação.**
6. Adicione um título à sua solicitação de pull e descreva todas as alterações que você está fazendo.
7. Envie a solicitação pull.

Um dos administradores do site processará sua solicitação de pull. Sua solicitação de pull será a tona no officeDev/site <repo name> em Problemas. Quando a solicitação pull for aceita, o problema será resolvido.

#### <a name="create-a-new-branch-after-merge"></a>Criar uma nova ramificação após a mesclagem

Depois que uma ramificação é mesclada com êxito (ou seja, sua solicitação pull é aceita), não continue trabalhando nesse branch local. Isso pode levar a conflitos de mesclagem se você enviar outra solicitação pull. Para fazer outra atualização, crie uma nova ramificação local a partir da ramificação upstream mesclada com êxito e exclua sua ramificação local inicial.

Por exemplo, se sua filial local X foi mesclada com êxito no branch mestre OfficeDev/microsoft-graph-docs e você deseja fazer atualizações adicionais para o conteúdo mesclado. Crie uma nova filial local, X2, a partir do branch mestre OfficeDev/microsoft-graph-docs. Para fazer isso, abra o GitBash e execute os seguintes comandos.

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Agora você tem cópias locais (em uma nova filial local) do trabalho enviado na ramificação X. A ramificação X2 também contém todos os trabalhos que outros autores mesclaram, portanto, se seu trabalho depender do trabalho de outras pessoas (por exemplo, imagens compartilhadas), ele estará disponível no novo branch. Você pode verificar se o trabalho anterior (e o trabalho de outras pessoas) está no branch fazendo check-out do novo branch...

    git checkout X2

... e verificando o conteúdo. (O `checkout` comando atualiza os arquivos `C:\Users\<yourusername>\microsoft-graph-docs` no estado atual da ramificação X2.) Depois de fazer check-out da nova filial, você pode fazer atualizações para o conteúdo e confirma-los como de costume. No entanto, para evitar trabalhar na ramificação mesclada (X) por engano, o melhor a fazer será excluí-la (confira a seguinte seção: **Excluir uma ramificação**).

#### <a name="delete-a-branch"></a>Excluir uma ramificação

Depois que suas alterações são mescladas com êxito no repositório principal, exclua a ramificação que você usou porque não precisa mais dele.  Qualquer trabalho adicional deve ser feito em um novo branch.  

#### <a name="to-delete-a-branch"></a>Para excluir um branch

1. No prompt de comando Git Bash, digite `git checkout master` . Isso garante que você não fique na ramificação a ser excluída (o que não é permitido).
2. Em seguida, no prompt de comando, digite `git branch -d <branch name>` . Isso exclui a ramificação em seu computador somente se tiver sido mesclada com êxito ao repositório upstream. (Você pode substituir esse comportamento com o sinalizador `–D`, mas primeiro certifique-se de que deseja fazer isso).
3. Por fim, digite `git push origin :<branch name>` no comando prompt (um espaço antes dos dois pontos e nenhum espaço depois deles).   Essa ação excluirá a ramificação em uma bifurcação do github.  

Parabéns, você contribuiu com êxito para o projeto!

## <a name="how-to-use-markdown-to-format-your-topic"></a>Como usar o Markdown para formatar seu tópico

### <a name="markdown"></a>Markdown

Todos os artigos neste repositório usam Markdown. Uma introdução completa (e listagem de toda a sintaxe) pode ser encontrada em [Daring Fireball - Markdown].

## <a name="faq"></a>Perguntas frequentes

### <a name="how-do-i-get-a-github-account"></a>Como obter uma conta do GitHub?

Preencha o formulário em [Ingressar no GitHub](https://github.com/join) para abrir uma conta gratuita do GitHub.

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Onde posso obter um Contrato de Licença do Colaborador?

Um aviso será automaticamente enviado para você informando que é preciso assinar o CLA (Contrato de Licença do Colaborador) se sua solicitação de recebimento exigir um.

Como membro da comunidade, **você deve assinar o CLA (Contrato de Licença do Colaborador) antes de poder contribuir com envios volumosos para esse projeto**. Você só precisa concluir e enviar a documentação uma vez. Reveja cuidadosamente o documento. Talvez seja necessário que seu empregador assine o documento.

### <a name="what-happens-with-my-contributions"></a>O que acontece com minhas contribuições?

Quando você enviar suas alterações, por meio de uma solicitação pull, nossa equipe será notificada e revisará sua solicitação de pull. Você receberá notificações sobre sua solicitação de pull do GitHub; você também pode ser notificado por alguém de nossa equipe se precisarmos de mais informações. Se sua solicitação de pull for aprovada, atualizaremos a documentação. Reservamo-nos o direito de editar seu envio para questões legais, de estilo, de clareza ou de outros problemas.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Posso me tornar um aprovador para as solicitações de GitHub pull deste repositório?

Atualmente, não estamos permitindo que colaboradores externos aprovem solicitações pull neste repositório.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Em quanto tempo receberei uma resposta sobre minha solicitação de alteração?

Normalmente, as solicitações de pull são revisadas dentro de 10 dias úteis.


## <a name="more-resources"></a>Mais recursos

* Para saber mais sobre Markdown, vá para o site do criador do [Markdown, Daring Fireball].
* Para saber mais sobre como usar o Git e GitHub, primeiro confira o [GitHub Ajuda].

[GitHub Home]: http://github.com
[Ajuda do GitHub]: http://help.github.com/
[Configurar o Git]: https://help.github.com/articles/set-up-git/
[Fireball audacioso - Markdown]: http://daringfireball.net/projects/markdown/
[Fireball audacioso]: http://daringfireball.net/
