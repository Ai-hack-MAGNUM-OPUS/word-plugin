import * as React from "react";
import { DefaultButton, ProgressIndicator } from "@fluentui/react";
import { load_docx } from "../../api/load_docx_text";
import { IEUID, IEData } from "../../api/interfaces";
import { update_state } from "../../api/update_state";
import ClipLoader from "react-spinners/ClipLoader";

/* global Word, require */


const processFunction = (context: Word.RequestContext, search_text: string, comment_text: string) => {
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
        ).getFirst().insertComment(comment_text);
    }
    return undefined
}


const App: React.FC = () => {

    const [comments, setComments] = React.useState<Word.Comment[]>([]);
    const [uid, setUid] = React.useState("");
    const [response_data, setResponseData] = React.useState<IEData>();

    React.useEffect(() => {
        if (uid.length) {
            setTimeout(async () => {
                const data: IEData = await update_state(uid);
                console.log(data)
                setResponseData(data);
            }, 2000)
            setUid("");
        }
        if (response_data != undefined) {
            Word.run(async (context) => {
                var comments = []
                const response_data_keys = Object.keys(response_data)
                for (var i = 0; i < response_data_keys.length; ++i) {
                    for (var j = 0; j < response_data[response_data_keys[i]].length; ++j) {
                        var comm = processFunction(context, response_data[response_data_keys[i]][j][0], response_data_keys[i]);
                        if (comm != undefined){
                            comments.push();
                        }
                    }
                }
                setComments(comments)
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
                    response_data === undefined && uid.length ?
                    <ClipLoader></ClipLoader> : <div></div>
                }
                </div>

        </div>
    );
}
export default App
// export default class App extends React.Component<AppProps, AppState> {
//   constructor(props, context) {
//     super(props, context);
//     this.state = {
//       listItems: [],
//     };
//   }

//   componentDidMount() {
//     this.setState({
//       listItems: [
//         {
//           icon: "Ribbon",
//           primaryText: "Achieve more with Office integration",
//         },
//         {
//           icon: "Unlock",
//           primaryText: "Unlock features and functionality",
//         },
//         {
//           icon: "Design",
//           primaryText: "Create and visualize like a pro",
//         },
//       ],
//     });
//   }

//   click = async () => {
//     return Word.run(async (context) => {
//       /**
//        * Insert your Word code here
//        */

//       // insert a paragraph at the end of the document.
//       const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//       // change the paragraph color to blue.
//       paragraph.font.color = "blue";

//       await context.sync();
//     });
//   };

//   render() {
//     const { title, isOfficeInitialized } = this.props;

//     if (!isOfficeInitialized) {
//       return (
//         <Progress
//           title={title}
//           logo={require("./../../../assets/logo-filled.png")}
//           message="Please sideload your addin to see app body."
//         />
//       );
//     }

//     return (
//       <div className="ms-welcome">
//         <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
//         <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
//           <p className="ms-font-l">
//             Modify the source files, then click <b>Run</b>.
//           </p>
//           <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
//             Run
//           </DefaultButton>
//         </HeroList>
//       </div>
//     );
//   }
// }
