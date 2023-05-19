import { loadingOFF, loadingON } from '../../loc/utils.js';
import { connectWithSPRest, getAuthorizedRequestOptionSP } from '../../loc/sharepoint.js';

async function init() {
  loadingON('Fetch token for sharepoint API\'s... please wait');
  await connectWithSPRest();
  loadingOFF('Token successful...done');
  const options = getAuthorizedRequestOptionSP();
  const listUrl = '/sites/adobecom/CC/www'; // dynamic
  const documentUrl = '/drafts/devashish/version-history.docx'; // dynamic
  document.getElementById('project-url').textContent = `${listUrl}${documentUrl}`;
  const url = `https://adobe.sharepoint.com/sites/adobecom/_api/web/GetFileByServerRelativeUrl('${listUrl}${documentUrl}')`;

  const fetchVersions = async (onlyMajorVersions = false) => {
    const documentData = await fetch(url, options);
    const {CheckInComment, TimeLastModified, UIVersionLabel} = await documentData.json();
  
    const currentVersion = {
      VersionLabel: UIVersionLabel,
      CheckInComment,
      Created: TimeLastModified
    }
  
    const versions = await fetch(`${url}/Versions`, options);
    const { value } = await versions.json();
  
    const versionHistory = [...value, currentVersion];
  
    const createTd = (data) => {
      const td = document.createElement('td');
      td.textContent = data;
      return td;
    }
  
    const createTr = (data) => {
      const trElement = document.createElement('tr');
      const { VersionLabel, CheckInComment, Created } = data;
      trElement.appendChild(createTd(VersionLabel));
      trElement.appendChild(createTd(Created.split('T')[0]));
      trElement.appendChild(createTd(CheckInComment));
      return trElement;
    }
    const versionDataParent = document.querySelector("#addVersionHistory");
    versionDataParent.innerHTML='';
    versionHistory.reverse().forEach((item) => {
      if(onlyMajorVersions) {
        if(item.VersionLabel.indexOf('.0') !== -1) {
          versionDataParent.appendChild(createTr(item));
        }
      } else {
        versionDataParent.appendChild(createTr(item));
      }
    });
  }

  const callOptions = getAuthorizedRequestOptionSP({
    method: 'POST'
  });

  const checkoutAPICall = async () => {
    await fetch(`${url}/CheckOut()`, callOptions);
  }

  const checkinAPICall = async (comment) => {
    await fetch(`${url}/CheckIn(comment='${comment}', checkintype='1')`, callOptions);
  }

  document.getElementById('update').addEventListener('click', async (e) => {
    e.preventDefault();
    loadingON('Creating new version');
    const comment = document.querySelector('#comment').value;
    if (comment) {
      await checkoutAPICall();
      await checkinAPICall(`Through API: ${comment}`);
      await fetchVersions();
    }
    loadingOFF('New version created');
  });

  const publishCommentCall = async (comment) => {
    const callOptions = getAuthorizedRequestOptionSP({
      method: 'POST'
    });
    await fetch(`${url}/Publish('${comment}')`, callOptions);
  }

  document.getElementById('publish').addEventListener('click', async (e) => {
    e.preventDefault();
    loadingON('Publish comment');
    const comment = document.querySelector('#comment').value;
    if (comment) {
      await publishCommentCall(comment);
      await fetchVersions();
    }
    loadingOFF('Published');
  });

  document.getElementById('majorVersions').addEventListener('click', async (e) => {
    e.preventDefault();
    loadingON('Fetch major Versions');
    await fetchVersions(true);
    loadingOFF('Major versions Listed');
  });

  document.getElementById('allVersions').addEventListener('click', async (e) => {
    e.preventDefault();
    loadingON('Fetch All Versions');
    await fetchVersions();
    loadingOFF('All Versions Listed');
  });

  fetchVersions();
}

export default init;
