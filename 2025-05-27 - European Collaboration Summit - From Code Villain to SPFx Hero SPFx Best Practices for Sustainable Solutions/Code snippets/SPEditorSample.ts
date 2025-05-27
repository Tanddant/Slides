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
  Title: string;
  "abstract": string;
  labels: string;
  startDate: Date;
  sessionGuid: string;
  sessionId: string
}

export interface ISpeaker {
  Title: string;
  Session: ISession;
}

const SpeakerSelects = ["Id", "Title", "Session/Id", "Session/labels", "Session/Title"];
const SpeakerExpands = [...SpeakerSelects.filter(s => s.indexOf("/") > -1).map(s => s.split("/")[0])];
const SessionSelects = ["Id", "Title", "abstract", "labels", "startDate", "sessionGuid", "sessionId"];


(async () => {
  try {












    //const query = sp.web.lists.filter("Title eq 'Sessions' or Title eq 'Speakers'").select("Id", "Title")
    //const query = sp.web.lists.filter(l => l.text("Title").equals("Sessions").or().text("Title").equals("Speakers")).select("Id", "Title")
    //const query = sp.web.lists.filter(l => l.text("Title").in("Sessions", "Speakers")).select("Id", "Title")





















    let Sessions: IItems = sp.web.lists.getByTitle("Sessions").items.select(...SessionSelects).top(1000);
    let Speakers: IItems = sp.web.lists.getByTitle("Speakers").items.select(...SpeakerSelects).expand(...SpeakerExpands).top(1000);



































    //  const query = Speakers.filter<ISpeaker>(speaker =>
    //   speaker.or(
    //     speaker.text("Title").equals("Dan Toft").and().lookup("Session").text("Title").contains("SPFx"),
    //     ...["SPFx","React","Development","M365"].map(subject => speaker.lookup("Session").text("Title").contains(subject))
    //   )
    // )

    //const query = Sessions.filter<ISession>(s => s.text("labels").contains("Development"));
    const query = Sessions.filter("substringof('Development', labels)");






    var url = new URL("https://localhost" + decodeURIComponent(query.toRequestUrl()))
    const filterstring = url.searchParams.get("$filter");
    const r = await query();
    bodyElement.innerHTML = `<pre style="margin:0"><code class="language-json">//${new Date().toISOString()}\n\n// ${r.length} result(s)\n\n// ${filterstring}\n\n${JSON.stringify(r, null, 4)}</code></pre>`;
    (window as any).hljs.highlightAll();
  } catch (exception) {
    bodyElement.innerHTML = exception;
  }
})().catch(console.log)