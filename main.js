// import { create, Whatsapp } from 'sulla';
const sulla = require('@open-wa/wa-automate');
var Firebird = require('node-firebird');
const util = require('util');
const os = require('os');
const fs = require('fs');

var db = null; // Banco de Dados

var botPhone = 'null';

function test(m_key) {
    console.log(m_key);
}

//https://github.com/danielcardeenas/sulla
//https://www.npmjs.com/package/node-firebird

console.log('Iniciando ... ');
abre_base();
function abre_base() {
	var SysKeys = getDBKeys();
    var options = {};
    options.database = __dirname + '\\dados.fdb';    
    options.user = SysKeys[0];
    options.password = SysKeys[1];
    options.lowercase_keys = false; // set to true to lowercase keys
    options.role = null;            // default
    options.pageSize = 4096;        // default when creating database
    //options.database = 'pra_50';

	console.log(os.hostname());
	
	options.host = '127.0.0.1';
	options.port = 19259;
	
	Firebird.attach(options, function(err, db2) {
		db = db2;
		if (err) {
			console.log("Ocorreu um erro Ao Abrir o Banco de Dados!");
			throw err;
		} else {
			console.log("Successo ao abrir o Banco de Dados!");
			sulla.create().then(client => start(client));
		}
	});
}

function getDBKeys() {
	var S = fs.readFileSync(__dirname + '\\inteligência\\Keys.txt', 'UTF-8');
	return S.split('|');
}

async function start(client) {
	var temp = await client.getMe();
	botPhone = temp.me._serialized
	console.log('Meu Numero: ' + botPhone);
	client.onMessage(message => {
		eventOnMessage(client,message);
	});
	handleIncomingMessages(client);
}

async function eventOnMessage(client,message) {
	
	var MSG_DATA = '';
	var MSG_CLI = '';
	var MSG_TYPE = 0;
	/* console.log(message);
	1 chat = Mensagem de Texto Normal - Texto no Body
	2 image = Imagem - Dados Brutos no Body
	3 ppt = Audio - Não Tem Body
	4 document = Arquivo - Body Vazio
	5 vcard = Contato
	*/
	if (message.type === 'chat') {
		MSG_TYPE = 1; // Chat
		MSG_CLI = message.from;
		MSG_DATA = message.body;
		console.log('<<< Recebido Texto "' + message.body + '" de ' + message.from);
	} else {
		return;
	}
	db.query('INSERT INTO TB_MSG (MSG_SEQ,MSG_CLI,MSG_BOT,MSG_DATA,MSG_TYPE,MSG_STATUS,MSG_DIR) VALUES (GEN_ID(GE_MSG_SEQ,1),?,?,?,?,0,0);', [MSG_CLI,botPhone,MSG_DATA,MSG_TYPE] , function(err, result) {
		if (err) {
			console.log("<<< Erro com o banco de dados ao tentar logar mensagen!");
			console.log(err);
		}
	});
}

async function handleIncomingMessages(client) {
	do {
		db.query('SELECT MSG_SEQ,MSG_DATA,MSG_CLI,MSG_TYPE,MSG_STATUS,MSG_DIR FROM TB_MSG WHERE MSG_DIR = 1 AND MSG_STATUS = 0 AND MSG_BOT = ? ORDER BY MSG_SEQ',[botPhone],function (err,result) {
			if (err) {
				console.log(err);
			} else if (result.length != 0) {
				for (var Z = 0;Z < result.length;Z++) {
					db.query('UPDATE TB_MSG SET MSG_STATUS = 1 WHERE MSG_SEQ = ?',[result[Z]["MSG_SEQ"]],function (err,result) {
						if (err) {
							console.log(">>> Erro ao marcar mensagem como enviada!");
							console.log(err);
						}	
					});
					if (result[Z]["MSG_TYPE"] == 1) { // Chat
						solveBlob(result[Z]["MSG_DATA"],result[Z]["MSG_CLI"],function (item,buffer) {
							console.log('>>> Eviando "' + buffer + '" para ' + item);
							client.sendText(item,buffer);
						});
					}
				}
			}
		});
		await sleep(1000);
	} while (true);
}

async function solveBlob(blob,param,callback) {
	var buffer = '';
	blob(function (err,name,e) {
		e.on('data', function(chunk) {
			buffer += chunk;
		});
		e.on('end', function() {
			callback(param,buffer);
		});
	});
}

/*

function verifica(client){
      // db = DATABASE
      // Faz a leitura de todas as mensagens 
	  console.log("Verificando...");
      var m_sql = 'SELECT W_SEQ, W_CELL, W_MESSAGE FROM TBWHATSAPP';
      m_sql += ' WHERE char_length(W_CELL) = 10 AND  W_SEND_DTH IS NULL AND (W_TIPO IS NULL OR W_TIPO = 0  OR W_TIPO = 1) AND W_CANCEL_DTH IS NULL';
	  m_sql += ' AND  W_ERROR_DTH IS NULL';
	  m_sql += ' ORDER BY W_SEQ';
      db.query(m_sql, function(err, result) {

          for (var i = 0, len = result.length; i < len; i++) {
            var m_row = result[i];
			console.log("-----------------------------------------");
			console.log("Enviando ..." + m_row['W_SEQ']);
            console.log("Enviando ..." + m_row['W_CELL']);
            console.log("Enviando ..." + m_row['W_MESSAGE']);
            //var res = client.sendText("55"+m_row['W_CELL'] + "@c.us", m_row['W_MESSAGE']);
			envia_txt(client,  m_row['W_SEQ'], m_row['W_CELL'] , m_row['W_MESSAGE'])
			
			//console.log(util.inspect(res, {showHidden: false, depth: null}))
			//console.log(util.inspect(res, false, null, true ))
			//console.log("Resultado :" + res);
			//console.log("Marcando como enviado");
			//console.log(client.checkNumberStatus("55"+m_row['W_CELL'] + "@c.us"));
			
            // Atualiza o status das mensagens
            //db.query('UPDATE TBWHATSAPP SET W_SEND_DTH = CURRENT_TIMESTAMP WHERE W_SEQ = ?;', [m_row['W_SEQ']], function(err, result) {
            
            //});            

          }          
        //  console.log(result[0]['W_MESSAGE']);
          //UPDATE TBWHATSAPP SET W_SEND_DTH = CURRENT_TIMESTAMP WHERE W_SEQ =

          // IMPORTANT: close the connection
          //db.detach();
      });

    
}

async function envia_txt(client, id, cell, txt) {
  var data = await client.sendText("55"+cell+ "@c.us", txt);
  console.log(id + ", " + data); // will print your data
  var data2 = data+' ';
  if (data2.indexOf('true') > -1) {
		db.query('UPDATE TBWHATSAPP SET W_SEND_DTH = CURRENT_TIMESTAMP WHERE W_SEQ = ?;', [id], function(err, result) {
			console.log('marcando ok');
        });            
		
  } else if (data2.indexOf('false') > -1) {
		db.query('UPDATE TBWHATSAPP SET W_ERROR_DTH = CURRENT_TIMESTAMP WHERE W_SEQ = ?;', [id], function(err, result) {
			console.log('marcando err');
        });            
		 
		client.sendText("558699983804@c.us",'Erro: *Mensagem não enviada ao contato:* ' + cell + "\n" + txt);
  } else {
		
  }
  
}
//verifica(client);

sulla.create().then(client => start(client));

function envia_menu(client, m_nome, m_from, body){
	console.log('ok-4');
			console.log(m_nome);
			console.log(m_nome.length);
			if (m_nome.length > 0 ){
				text = "Bem Vindo, " + m_nome ;
				client.sendText(m_from, text);			
			}
			text = "*Informativo* \n _Essa conta é usada somente para envio de informações_";
			client.sendText(m_from, text);

			text = "Especifique o setor no qual quer entrar em contato: \n";
			text += "1 - Pax 24h\n";
			text += "2 - Controladoria\n";
			text += "3 - Setor de TI\n";
			if (m_nome.length > 0 ){
				text += "8 - Abrir Chamado no TI\n";
			}
			text += "9 - Outras funções";
			client.sendText(m_from, text);

			//  client.sendText("558699983804@c.us", client.getContact(message.from));
			if (m_from!=='558699983804@c.us'){
			  client.sendText("558699983804@c.us", m_from + "\n" + body);
			  client.sendText("558694227127@c.us", m_from + "\n" + body);
			}
}
function start(client) {
  var text="";
  console.log('Iniciando 2 ... \n');
  client.sendText("558699983804@c.us", "Iniciando Servidor Java");
  abre_base();
  console.log('Base aberta ...1 \n');
    //
    intervalid = setInterval(verifica, 3000, client);
 console.log('Base aberta ...2 \n');
    client.onMessage(async message => {
      if (message.body === 'Hi') {
        client.sendText(message.from, '?? Hello from sulla!');
      } else if (message.body === '1') {
			//Pax 24h
			try {
				client.sendContact(message.from,"558694183880@c.us")
			} catch (err) {
				console.log(err.message);
				client.sendText("558699983804@c.us",err.message);
			}    
      } else if (message.body === '2') {
			//Controladoria
			client.sendContact(message.from,"558694184663@c.us")

      } else if (message.body === '3') {
			//TI
			client.sendContact(message.from,"558699983804@c.us")
		} else if (message.body === '8') {
			client.sendText(message.from, 'Você não tem acesso a esse menu, caso deseje acesso a essa função entre em contato com o setor de TI');
	
      } else if (message.body === '9') {
        client.sendText(message.from, 'Você não tem acesso a esse menu, caso deseje acesso a essas funções entre em contato com o setor de TI');

      } else {
			var m_numero = message.from.substr(2,10)+' ';			
			var m_sql = 'SELECT USER_NOME FROM USERS WHERE (select oparam from PROC_GET_ONLY_NUMBERS( USER_WHATSAPP)) = ?';
			db.query(m_sql, [m_numero], function(err, result) {
			  var m_nome = '';
			  if (result.length>0){
				var m_row = result[0];
				m_nome = m_row['USER_NOME'];
			  }
			  envia_menu(client, m_nome, message.from, message.body)
			});
			console.log("--------------------------------------------");			
			console.log('from: ' + message.from);
			console.log('type: ' + message.type);
			console.log('isLink: ' + message.isLink);
			console.log('isMedia: ' + message.isMedia);			
			console.log('');
      }    
    });

}
*/
function sleep(ms) {
	return new Promise(resolve => setTimeout(resolve, ms));
}
