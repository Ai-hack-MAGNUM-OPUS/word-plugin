import * as React from "react";
import { DefaultButton, ProgressIndicator } from "@fluentui/react";
import { load_docx } from "../../api/load_docx_text";
import { IEUID, IEData } from "../../api/interfaces";
import { update_state } from "../../api/update_state";
import ClipLoader from "react-spinners/ClipLoader";

/* global Word, require */


const processFunction = (context: Word.RequestContext, search_text: string, comment_text: string, score: number) => {
    //if (search_text.includes("\u")) return undefined;
    search_text.replace("\u0005", "")
    console.log(search_text)
    if (search_text.length > 10){
        return context.document.body.search(
            search_text.slice(0, 255).split("\\").join(""),
            {
                ignorePunct: true,
                ignoreSpace: true
            }
        ).getFirst().insertComment(comment_text+ "\n" + "Оценка: " + score.toString());
    }
    return undefined
}


const App: React.FC = () => {

    const [comments, setComments] = React.useState<Word.Comment[]>([]);
    const [request_pull, setRequestPull] = React.useState(false)
    const [uid, setUid] = React.useState("");
    const [not_defined_fields, setNotDefinedFields] = React.useState<string[]>([]);
    const [response_data, setResponseData] = React.useState<IEData>(undefined);

    React.useEffect(() => {
        if (uid.length) {
            setTimeout(async () => {
                const data: IEData = await update_state(uid);
                console.log(data)
                setResponseData(data);
                setRequestPull(false);
            }, 10000)
            setUid("");
        }
        if (response_data != undefined) {
            Word.run(async (context) => {
                var comments = []
                var not_defined_fields = []
                const response_data_keys = Object.keys(response_data)
                for (var i = 0; i < response_data_keys.length; ++i) {
                    if (!response_data[response_data_keys[i]].length) {
                        not_defined_fields.push(response_data_keys[i])
                    }
                    for (var j = 0; j < response_data[response_data_keys[i]].length; ++j) {
                        var comm = processFunction(
                            context, 
                            response_data[response_data_keys[i]][j][0], 
                            response_data_keys[i],
                            response_data[response_data_keys[i]][j][1]
                        );
                        if (comm != undefined){
                            comments.push(comm);
                        }
                    }
                }
                setComments(comments)
                setNotDefinedFields(not_defined_fields);
                setResponseData(undefined)
                //
            })
            setResponseData(undefined);
        }
    })
    return (
        <div style={{
            'display': 'flex',
            'justifyContent': 'center',
            'alignContent': 'center',
            'flexDirection': 'column'
        }}>
            <DefaultButton
                    onClick={() => {
                        Word.run(function(context) {
                            // Insert your code here. For example:
                            var documentBody = context.document.body;
                            context.load(documentBody);
                            return context.sync()
                            .then( async () => {
                                const data: IEUID = await load_docx(documentBody.text);
                                console.log(data);
                                setUid(data.uuid);
                                setRequestPull(true);
                            })
                        });
                    }}
                >
                    Проверить на ошибки
                </DefaultButton>
                <div style={{
                    'display': 'flex',
                    'justifyContent': 'center',
                    'alignItems': 'center',
                    'marginTop': 100
                }}>
                {
                    response_data == undefined && request_pull ?
                    <ClipLoader></ClipLoader> : <div></div>
                }
                </div>
                <ul>
                    {
                        not_defined_fields.map((e) => {
                            <li>{e}</li>
                        })
                    }
                </ul>

        </div>
    );
}
export default App
