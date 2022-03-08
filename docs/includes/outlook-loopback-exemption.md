> [!NOTE]
> Se você estiver executando o seu complemento no localhost e vir o erro "Lamentamos, não foi possível acessar *{your-add-in-name-here}*. Certifique-se de ter uma conexão de rede. Se o problema continuar, tente novamente mais tarde.", talvez seja necessário habilitar uma isenção de loopback.
>
> 1. Close Outlook.
> 1. Abra o **Gerenciador de Tarefas** e certifique-se de que o **msoadfsb.exe** não está em execução.
> 1. De definir [a isenção de loopback](/previous-versions/windows/apps/hh780593(v=win.10)?redirectedfrom=MSDN) em um prompt elevado.
>     - Se você estiver usando e porta `https://localhost` 3000 (a configuração padrão), execute o seguinte comando.
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_https___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>     - Se você estiver usando e porta `http://localhost` 3000, execute o seguinte comando.
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>
>      **Observação**: se você não estiver usando a porta padrão 3000, substitua-a no comando pelo número de porta real.
> 1. Reinicie o Outlook.
