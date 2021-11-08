Se seu projeto for baseado em node.js (ou seja, não desenvolvido com o Visual Studio e o servidor de Informações da Internet (IIS),você poderá forçar o Office no Windows a usar o Edge Legacy ou o Internet Explorer para executar os complementos, mesmo que você tenha uma combinação de versões Windows e Office que normalmente usariam um navegador mais recente. Para obter mais informações sobre quais navegadores são usados por várias combinações de versões Windows e Office, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

1. Se seu projeto não *foi* criado com a ferramenta Yo Office, você precisará instalar a ferramenta office-addin-dev-settings. Execute o seguinte comando em um prompt de comando.

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. Especifique o navegador que você Office usar com o seguinte comando em um prompt de comando na raiz do projeto. Substitua pelo caminho relativo, que é apenas o nome do arquivo de manifesto se `<path-to-manifest>` estiver na raiz do projeto. Substitua `<webview>` por um ou `ie` `edge-legacy` .

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    Apresentamos um exemplo a seguir.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    Você deve ver uma mensagem na linha de comando que o tipo de webview agora está definido como IE (ou Edge Legacy).

1. Quando terminar, desem Office para continuar usando o navegador padrão para sua combinação de versões Windows e Office com o comando a seguir.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
