import * as React from "react";
import { DefaultButton } from "@fluentui/react";

/* global Word, require */


const App: React.FC = () => {

    const [process, onProcessed] = React.useState(false);

    return (
        <div style={{
            'display': 'flex',
            'justifyContent': 'center',
            'alignContent': 'center'
        }}>
            <DefaultButton
                onClick={() => {
                    Word.run(function(context) {
                        // Insert your code here. For example:
                        var documentBody = context.document.body;
                        context.load(documentBody);
                        return context.sync()
                        .then(function(){
                            console.log(documentBody.text);
                        })
                    });
                    Word.run(async (context) => {
                        Office.context.document.getFileAsync(Office.FileType.Compressed, (file) => {
                            console.log(file.value);
                        })
                        context.document.body.search("2-кратного размера").getFirst().insertComment("fuck you")
                    })
                }}
            >
                Проверить на ошибки
            </DefaultButton>
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
