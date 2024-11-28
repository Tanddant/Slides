import { spfi, SPBrowser, IItems } from "@pnp/sp/presets/all";
import { InjectHeaders } from "@pnp/queryable";

const elem_Id = "PnP_Rocks"
let bodyElement: HTMLDivElement = document.getElementById(elem_Id) as HTMLDivElement
if (!bodyElement) {
  let elem = document.createElement("div")
  elem.id = elem_Id
  elem.style.height = "100vh"
  elem.style.width = "100vw"
  elem.style.zIndex = "10000"
  elem.style.position = "absolute"
  elem.style.backgroundColor = "white"
  elem.style.top = "0"
  elem.style.overflow = "scroll"

  bodyElement = document.body.appendChild(elem);
}

const HightlightJS_ID = "HighlightJSscriptLink"
let HighlightJs: HTMLScriptElement = document.getElementById(HightlightJS_ID) as HTMLScriptElement;
if (!HighlightJs) {
  let elem = document.createElement("script");
  elem.src = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js";
  document.head.appendChild(elem);
}

const HighlightCSS_Id = "HightlightCSS"
let HighlightCSS: HTMLLinkElement = document.getElementById(HighlightCSS_Id) as HTMLLinkElement;
if (!HighlightCSS) {
  let elem = document.createElement("link");
  elem.rel = "stylesheet"
  elem.href = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/atom-one-dark-reasonable.min.css"
  document.head.append(elem)
}

const sp = spfi().using(
  SPBrowser({ baseUrl: (window as any)._spPageContextInfo.webAbsoluteUrl }),
  InjectHeaders({ Accept: 'application/json; odata=nometadata' })
);


export interface ISpeaker {
  Name: string;
  Company: string;
  SessionCount: number;
}

export interface ISession {
  Speaker: ISpeaker;
  SessionCode: string;
  Title: string;
  Category: string;
  SessionLevel: number;
}

const SpeakerSelects = ["Id", "Name", "Company", "SessionCount"];
const SessionSelects = ["Id", "SessionCode", "Title", "Category", "Title", "SessionLevel", "Speaker/Name", "Speaker/Company", "Speaker/SessionCount", "Speaker/Id"];
const SessionExpands = [...SessionSelects.filter(s => s.indexOf("/") > -1).map(s => s.split("/")[0])];


(async () => {


  let Sessions: IItems = sp.web.lists.getByTitle("Sessions").items.select(...SessionSelects).expand(...SessionExpands).top(1000);
  let Speakers: IItems = sp.web.lists.getByTitle("Speakers").items.select(...SpeakerSelects).top(1000);


  //const query = Sessions.filter<ISession>(s => s.text("Category").equals("Microsoft Teams"));
  const query = Sessions.filter("Category eq 'Microsoft Teams'");







  const filterstring = new URLSearchParams(query.toRequestUrl()).get("$filter");
  const r = await query();

  bodyElement.innerHTML = `<pre style="margin:0"><code class="language-json">// ${r.length} result(s)\n\n// ${filterstring}\n\n${JSON.stringify(r, null, 4)}</code></pre>`;
  (window as any).hljs.highlightAll();
})().catch(console.log)








