# Acompanhamento de obra

Painel local de acompanhamento do empreendimento BoulevarDiálogo Butantã, com:

- controle por torre, pavimento, apartamento e serviço;
- ocorrências com fotos e histórico em PDF;
- dashboards e evolução visual dos pavimentos;
- relatórios e QR Codes para visitantes;
- persistência em SQLite;
- acesso administrativo protegido por login.

## Instalação

É necessário ter Python 3 instalado.

```bash
python3 -m pip install -r requirements.txt
cp configuracao.exemplo.json configuracao.local.json
```

Edite `configuracao.local.json` e defina o usuário e a senha do engenheiro. Esse arquivo e o banco SQLite são ignorados pelo Git.

## Executar

No macOS, abra `INICIAR-PAINEL.command`, ou execute:

```bash
python3 servidor.py
```

No computador, acesse `http://127.0.0.1:8000`.

Para acesso por celular, use o endereço de rede mostrado no terminal e mantenha o celular na mesma rede Wi-Fi.

## Dados locais

O arquivo `acompanhamento.db` é criado automaticamente e não deve ser publicado, pois pode conter dados e fotos da obra.

