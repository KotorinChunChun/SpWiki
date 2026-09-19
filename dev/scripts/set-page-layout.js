/*
 * Switches a modern page between the normal article layout and the single part
 * app page layout, without PnP PowerShell or CLI for Microsoft 365.
 *
 * `Set-PnPPage -LayoutType SingleWebPartAppPage` does exactly one thing under
 * the hood: it writes the PageLayoutType field of the page's Site Pages list
 * item. That is a plain REST call, so it can be done from the browser using the
 * session you are already signed in with — no module install, no Entra app
 * registration, no admin consent.
 *
 * HOW TO USE
 *   1. Create the page normally, add the SpWiki web part, configure it
 *      and publish. (Once the page is a single part app page it can no longer be
 *      edited in the browser, so do this first.)
 *   2. Open any page of the same site in Edge or Chrome.
 *   3. F12 -> Console. If the console asks you to, type: allow pasting
 *   4. Edit the three constants below, paste the whole file, press Enter.
 *
 * To undo, set LAYOUT to 'Article' and run it again.
 */
(async () => {
  // ---------------------------------------------------------------- settings
  const WEB_URL = '/sites/docs';        // server relative site URL
  const PAGE = 'md2.aspx';                      // page file name in SitePages
  const LAYOUT = 'SingleWebPartAppPage';        // or 'Article' to undo
  // -------------------------------------------------------------------------

  const web = WEB_URL.replace(/\/+$/, '');
  const pagePath = `${web}/SitePages/${PAGE}`;
  const json = { Accept: 'application/json;odata=nometadata' };

  const read = async (url) => {
    const response = await fetch(url, { headers: json, credentials: 'include' });
    if (!response.ok) {
      throw new Error(`${response.status} ${response.statusText} for ${url}`);
    }
    return response.json();
  };

  try {
    const context = await (
      await fetch(`${web}/_api/contextinfo`, { method: 'POST', headers: json, credentials: 'include' })
    ).json();

    const before = await read(
      `${web}/_api/web/GetFileByServerRelativeUrl('${pagePath}')/ListItemAllFields`
    );
    console.log(`current layout: ${before.PageLayoutType}`);

    const response = await fetch(
      `${web}/_api/web/GetFileByServerRelativeUrl('${pagePath}')/ListItemAllFields`,
      {
        method: 'POST',
        credentials: 'include',
        headers: {
          ...json,
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
          'X-RequestDigest': context.FormDigestValue
        },
        body: JSON.stringify({ PageLayoutType: LAYOUT })
      }
    );

    if (!response.ok) {
      console.error(`update failed: ${response.status} ${response.statusText}`);
      console.error(await response.text());
      return;
    }

    const after = await read(
      `${web}/_api/web/GetFileByServerRelativeUrl('${pagePath}')/ListItemAllFields`
    );
    console.log(`new layout    : ${after.PageLayoutType}`);
    console.log(`open: ${location.origin}${pagePath}?file=README.md`);
  } catch (error) {
    console.error(error);
  }
})();
