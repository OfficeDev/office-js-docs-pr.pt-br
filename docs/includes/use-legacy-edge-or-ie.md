Se seu projeto for baseado em node.js (ou seja, não desenvolvido com o Visual Studio e o IIS (Servidor de Informações da Internet),você poderá forçar o Office no Windows a usar o Edge Legacy ou o Internet Explorer para executar suplementos, mesmo que você tenha uma combinação de versões do Windows e do Office que normalmente usariam um navegador mais recente. Para obter mais informações sobre quais navegadores são usados por várias combinações de versões do Windows e do Office, consulte Navegadores usados [pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!NOTE]
> A ferramenta usada para forçar a alteração no navegador só tem suporte no canal de assinatura Beta do Microsoft 365. Ingresse [no programa Office Insider](https://insider.office.com/join/windows) e selecione a **opção Canal Beta** para acessar builds do Office Beta. Consulte também [Sobre o Office: Qual versão do Office estou usando?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).
>
> Estritamente, é a opção `webview` dessa ferramenta (consulte a **Etapa 2**) que requer o canal Beta. A ferramenta tem outras opções que não têm esse requisito.

1. Se o projeto não *tiver* sido criado com a ferramenta [yeoman para suplementos do Office](../develop/yeoman-generator-overview.md) , você precisará instalar a ferramenta office-addin-dev-settings. Execute o comando a seguir em um prompt de comando.

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. Especifique o navegador que você deseja que o Office use com o comando a seguir em um prompt de comando na raiz do projeto. Substitua `<path-to-manifest>` pelo caminho relativo, que é apenas o nome do arquivo de manifesto se ele estiver na raiz do projeto. Substitua `<webview>` por um ou `edge-legacy``ie` .

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    Apresentamos um exemplo a seguir.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    Você deverá ver uma mensagem na linha de comando informando que o tipo de modo de exibição da Web agora está definido como IE (ou Edge Legacy).

1. Quando terminar, defina o Office para retomar usando o navegador padrão para sua combinação de versões do Windows e do Office com o comando a seguir.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
