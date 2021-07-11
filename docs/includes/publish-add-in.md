Um Suplemento do Office consiste em um aplicativo Web e um arquivo de manifesto. O aplicativo Web define a interface do usuário e a funcionalidade do suplemento, enquanto o manifesto especifica o local do aplicativo Web e define as configurações e os recursos do suplemento. 

Enquanto estiver desenvolvendo o suplemento, você poderá executá-lo em seu servidor Web local (`localhost`), mas quando estiver pronto para publicá-lo para que outros usuários acessem o, será necessário implantar o aplicativo Web em um servidor Web ou serviço de hospedagem na Web (por exemplo, Microsoft Azure) e atualizar o manifesto para especificar a URL do aplicativo implantado. 

Quando o suplemento estiver funcionando conforme desejado e você estiver pronto para publicá-lo para que outros usuários acessem, conclua as etapas a seguir.

1. Na linha de comando, no diretório raiz do projeto de suplemento, execute o comando a seguir para preparar todos os arquivos para implantação de produção.

    ```command&nbsp;line
    npm run build
    ```

    Quando a compilação for concluída, a pasta **dist** no diretório raiz do projeto de suplemento incluirá os arquivos que você implantará nas etapas subsequentes.

2. Carregue o conteúdo da pasta **dist** para o servidor Web que hospedará o suplemento. Você pode usar qualquer tipo de servidor Web ou serviço de hospedagem na Web para hospedar o suplemento.

3. No VS Code, abra o arquivo de manifesto do suplemento localizado na pasta raiz do projeto (`manifest.xml`). Substitua todas as ocorrências de `https://localhost:3000` pela URL do aplicativo Web implementado no servidor Web na etapa anterior.

4. Escolha o método que deseja usar para [implantar seu suplemento do Office](../publish/publish.md) e siga as instruções para publicar o arquivo de manifesto.
