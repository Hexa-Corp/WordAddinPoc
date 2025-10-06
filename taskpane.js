
/* global Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office is ready!");
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    document.getElementById("run").addEventListener("click", showSharePointDataRaw);
  }
});

async function showSharePointDataRaw() {
  try {
    const tokenRes = await fetch("http://localhost:3001/getAppToken");
    const tokenData = await tokenRes.json();
    const token = tokenData.token;

    const listUrl = "https://graph.microsoft.com/v1.0/sites/tylky.sharepoint.com,76aa5505-18b0-420c-aac0-fb41b3f15a31,304365ac-39cb-4017-9892-b2642829d467/lists/0f565f3b-e4be-4dd9-ad16-9803cbee809f/items?expand=fields";
    const res = await fetch(listUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json"
      }
    });

    if (!res.ok) throw new Error(`Graph API error: ${res.status} ${res.statusText}`);

    const data = await res.json();

    await Word.run(async (context) => {
      const docBody = context.document.body;

      docBody.clear();

      const jsonString = JSON.stringify(data.value, null, 2);
      docBody.insertParagraph(jsonString, Word.InsertLocation.start);

      await context.sync();
    });

  } catch (err) {
    console.error("Error fetching SharePoint list:", err);
    // alert("Failed to load SharePoint list. See console for details.");
  }
}
