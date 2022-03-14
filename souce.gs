var planilha = SpreadsheetApp.getActive();
var guiaDados = planilha.getSheetByName("Pipeline Variation");
var guiaMenu = planilha.getSheetByName("Cadastro Pipeline");

var date = new Date();
var month = date.getMonth() + 1;

var data = guiaDados.getRange(2, 1, guiaDados.getLastRow(), 25).getValues();
var head = guiaMenu.getRange("B1:B22").getValues();

function Search() {
    if (head[0] != "") {
        for (var lin = 0; lin < data.length; lin++) {
            if (data[lin][0] == head[0] && data[lin][22] == month) {
                guiaMenu.getRange(1, 2).setValue(data[lin][0]); //customer_id
                guiaMenu.getRange(3, 2).setValue(data[lin][2]); //csm
                guiaMenu.getRange(5, 2).setValue(data[lin][4]); //ticket
                guiaMenu.getRange(6, 2).setValue(data[lin][5]); //solution
                guiaMenu.getRange(9, 2).setValue(data[lin][8]); //MRR
                guiaMenu.getRange(11, 2).setValue(data[lin][10]); //type
                guiaMenu.getRange(12, 2).setValue(data[lin][11]); //initValue
                guiaMenu.getRange(13, 2).setValue(data[lin][12]); //realValue
                guiaMenu.getRange(14, 2).setValue(data[lin][13]); //wrinkle
                guiaMenu.getRange(15, 2).setValue(data[lin][14]); //status
                guiaMenu.getRange(16, 2).setValue(data[lin][15]); //initDate
                guiaMenu.getRange(17, 2).setValue(data[lin][16]); //expectationDate
                guiaMenu.getRange(18, 2).setValue(data[lin][17]); //endDate
                guiaMenu.getRange(19, 2).setValue(data[lin][18]); //paymentType
                guiaMenu.getRange(20, 2).setValue(data[lin][19]); //motivation
                guiaMenu.getRange(21, 2).setValue(data[lin][20]); //obs
                guiaMenu.getRange(22, 2).setValue(data[lin][24]); //id

                Browser.msgBox("🙌 Sucesso", "Registro encontrato.", Browser.Buttons.OK);
            }
            /*else {
                     Browser.msgBox("⛔ Alerta", "Registro não encontrado", Browser.Buttons.OK);
                 }*/
        }
    } else {
        Browser.msgBox("⛔ Alerta", "Para concluir pesquisa, preencha o campo \"Id Heimdall\". ✏️", Browser.Buttons.OK);
    }
}

function Clear() {
    guiaMenu.getRangeList(['B1', 'B3', 'B5:B6', 'B9', 'B11:B22']).clear({
        contentsOnly: true,
        kipFilteredRows: true
    });
}

function Salvar() {
    if (head[21] != "") {
        //Edit
        for (var lin = 0; lin < data.length; lin++) {
            if (data[lin][0] == head[0] && data[lin][22] == month || data[lin][24] == head[21]) {
                var confirmation = Browser.msgBox("⚠️ Registro já existe!!!", "Este cliente já contém um registro neste mês. Deseja sobrescrever❓", Browser.Buttons.YES_NO);

                if (confirmation == "yes") {
                    if (head[0] == "" || head[2] == "" || head[5] == "" || head[8] == "" || head[10] == "" || head[14] == "" || head[16] == "") {
                        Browser.msgBox("⛔ Alerta", "Campos obrigatórios em branco", Browser.Buttons.OK);
                    } else {
                        var row = lin + 2;

                        guiaDados.getRange(row, 1).setValue(head[0]);
                        guiaDados.getRange(row, 2).setValue(head[1]);
                        guiaDados.getRange(row, 3).setValue(head[2]);
                        guiaDados.getRange(row, 4).setValue(head[3]);
                        guiaDados.getRange(row, 5).setValue(head[4]);
                        guiaDados.getRange(row, 6).setValue(head[5]);
                        guiaDados.getRange(row, 7).setValue(head[6]);
                        guiaDados.getRange(row, 8).setValue(head[7]);
                        guiaDados.getRange(row, 9).setValue(head[8]);
                        guiaDados.getRange(row, 10).setValue(head[9]);
                        guiaDados.getRange(row, 11).setValue(head[10]);
                        guiaDados.getRange(row, 12).setValue(head[11]);
                        guiaDados.getRange(row, 13).setValue(head[12]);
                        guiaDados.getRange(row, 14).setValue(head[13]);
                        guiaDados.getRange(row, 15).setValue(head[14]);
                        guiaDados.getRange(row, 16).setValue(head[15]);
                        guiaDados.getRange(row, 17).setValue(head[16]);
                        guiaDados.getRange(row, 18).setValue(head[17]);
                        guiaDados.getRange(row, 19).setValue(head[18]);
                        guiaDados.getRange(row, 20).setValue(head[19]);
                        guiaDados.getRange(row, 21).setValue(head[20]);

                        Browser.msgBox("🙌 Sucesso", "Registro editado com sucesso", Browser.Buttons.OK);

                        Clear();
                    }
                } else {
                    guiaMenu.getRange(1, 2).setValue(data[lin][0]); //customer_id
                    guiaMenu.getRange(3, 2).setValue(data[lin][2]); //csm
                    guiaMenu.getRange(5, 2).setValue(data[lin][4]); //ticket
                    guiaMenu.getRange(6, 2).setValue(data[lin][5]); //solution
                    guiaMenu.getRange(9, 2).setValue(data[lin][8]); //MRR
                    guiaMenu.getRange(11, 2).setValue(data[lin][10]); //type
                    guiaMenu.getRange(12, 2).setValue(data[lin][11]); //initValue
                    guiaMenu.getRange(13, 2).setValue(data[lin][12]); //realValue
                    guiaMenu.getRange(14, 2).setValue(data[lin][13]); //wrinkle
                    guiaMenu.getRange(15, 2).setValue(data[lin][14]); //status
                    guiaMenu.getRange(16, 2).setValue(data[lin][15]); //initDate
                    guiaMenu.getRange(17, 2).setValue(data[lin][16]); //expectationDate
                    guiaMenu.getRange(18, 2).setValue(data[lin][17]); //endDate
                    guiaMenu.getRange(19, 2).setValue(data[lin][18]); //paymentType
                    guiaMenu.getRange(20, 2).setValue(data[lin][19]); //motivation
                    guiaMenu.getRange(21, 2).setValue(data[lin][20]); //obs
                    guiaMenu.getRange(22, 2).setValue(data[lin][24]); //id
                }
            }
        }
    } else {
        //NEW
        for (var lin = 0; lin < data.length; lin++) {
            // Browser.msgBox(data[lin][24]);
            if (data[lin][0] != head[0] && data[lin][22] != month && data[lin][24] != head[21]) {
                if (head[0] == "" || head[2] == "" || head[5] == "" || head[8] == "" || head[10] == "" || head[14] == "" || head[16] == "") {
                    Browser.msgBox("⛔ Alerta", "Campos obrigatórios em branco", Browser.Buttons.OK);
                } else {
                    var newId = Math.max.apply(null, guiaDados.getRange("Y2:Y").getValues());
                    var newId = newId + 1;
                    var lin = guiaDados.getLastRow() + 1;
                    // Browser.msgBox("teste", lin, Browser.Buttons.OK_CANCEL);

                    guiaDados.getRange(lin, 1).setValue(head[0]); //customer_id 
                    guiaDados.getRange(lin, 2).setValue(head[1]); //customer_name
                    guiaDados.getRange(lin, 3).setValue(head[2]); //csm
                    guiaDados.getRange(lin, 4).setValue(head[3]); //team
                    guiaDados.getRange(lin, 5).setValue(head[4]); //ticket
                    guiaDados.getRange(lin, 6).setValue(head[5]); //solution
                    guiaDados.getRange(lin, 7).setValue(head[6]); //country
                    guiaDados.getRange(lin, 8).setValue(head[7]); //group
                    guiaDados.getRange(lin, 9).setValue(head[8]); //mrr
                    guiaDados.getRange(lin, 10).setValue(head[9]); //lt
                    guiaDados.getRange(lin, 11).setValue(head[10]); //type
                    guiaDados.getRange(lin, 12).setValue(head[11]); //initValue
                    guiaDados.getRange(lin, 13).setValue(head[12]); //realValue
                    guiaDados.getRange(lin, 14).setValue(head[13]); //wrinkle
                    guiaDados.getRange(lin, 15).setValue(head[14]); //status
                    guiaDados.getRange(lin, 16).setValue(head[15]); //initDate
                    guiaDados.getRange(lin, 17).setValue(head[16]); //expectionDate
                    guiaDados.getRange(lin, 18).setValue(head[17]); //endDate
                    guiaDados.getRange(lin, 19).setValue(head[18]); //paymentType
                    guiaDados.getRange(lin, 20).setValue(head[19]); //motivation
                    guiaDados.getRange(lin, 21).setValue(head[20]); //obs
                    guiaDados.getRange(lin, 23).setValue(month); //month
                    guiaDados.getRange(lin, 25).setValue(newId); //id

                    Browser.msgBox("🙌 Sucesso", "Registro cadastrado com sucesso", Browser.Buttons.OK);

                    Clear();
                }
            }
        }
    }
}
