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


export interface ISession {
  Id: number,
  Code: string,
  Title: string,
  Audience: string,
  Level: string,
  Track: string,
  Speaker: ISpeaker,
}

export interface ISpeaker {
  Id: number;
  Title: string;
  Company: string;
  Country: string
}

const SessionSelects = ["Id", "Code", "Title", "Audience", "Level", "Track", "Speaker/Id", "Speaker/Title", "Speaker/Company", "Speaker/Country"];
const SessionsExpands = [...SessionSelects.filter(s => s.indexOf("/") > -1).map(s => s.split("/")[0])];
const SpeakerSelects = ["Id", "Title", "Company", "Country"];


(async () => {
  try {





    let Sessions: IItems = sp.web.lists.getByTitle("Sessions").items.select(...SessionSelects).expand(...SessionsExpands).top(1000);
    let Speakers: IItems = sp.web.lists.getByTitle("Speakers").items.select(...SpeakerSelects).top(1000);







    // const query = Speakers.filter<ISpeaker>(s => s.text("Country").equals("Ireland"))
    // const query = Speakers.filter("Country eq 'Ireland'")














    //const query = Sessions.filter<ISession>(s => s.lookup("Speaker").text("Country").equals("Ireland"))










    // const topicsImInterestedIn: string[] = ["Development", "SPFx", "C#", "API"]
    // const query = Sessions.filter<ISession>(f =>
    //   f.and(
    //     f.or(
    //       ...topicsImInterestedIn.map(topic => f.text("Title").contains(topic))
    //     ),
    //     f.lookup("Speaker").text("Title").startsWith("d")
    //   )
    // )

















    // const query = Sessions.filter<ISession>(s => s.text("Title").contains("SPFx"));
    // const query = Sessions.filter<ISession>(s => s.lookup("Speaker").text("Title").contains("Dan"));
    // const query = Sessions.filter("substringof('Development', labels)");






    var url = new URL("https://localhost" + decodeURI(query.toRequestUrl()))
    const filterstring = url.searchParams.get("$filter");
    console.log(query.toRequestUrl())
    const r = await query();
    bodyElement.innerHTML = `<pre style="margin:0"><code class="language-json">//${new Date().toISOString()}\n\n// ${r.length} result(s)\n\n// ${filterstring}\n\n${JSON.stringify(r, null, 4)}</code></pre>`;
    (window as any).hljs.highlightAll();
  } catch (exception) {
    bodyElement.innerHTML = exception;
  }
})().catch(console.log)