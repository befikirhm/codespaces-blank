// app.js
const { useState, useEffect } = React;

const getDigest = () => {
  return $.ajax({
    type: 'POST',
    url: `${_spPageContextInfo.webAbsoluteUrl}/_api/contextinfo`,
    headers: {
      "Accept": "application/json; odata=verbose"
    },
    xhrFields: { withCredentials: true }
  }).then(data => data.d.GetContextWebInformation.FormDigestValue)
    .catch(error => {
      console.error('Error fetching digest:', error);
      throw new Error('Failed to fetch request digest token.');
    });
};

const App = () => {
  const [surveys, setSurveys] = useState([]);
  const [userRole, setUserRole] = useState('');
  const [currentUser, setCurrentUser] = useState(null);
  const [isSiteAdmin, setIsSiteAdmin] = useState(false);
  const [filters, setFilters] = useState({ status: [], search: '' });
  const [isSideNavOpen, setIsSideNavOpen] = useState(false);
  const [isLoadingUser, setIsLoadingUser] = useState(true);
  const [isLoadingSurveys, setIsLoadingSurveys] = useState(false);
  const [notifications, setNotifications] = useState([]);

  const addNotification = (message, type = 'success') => {
    const id = Date.now();
    setNotifications(prev => [...prev, { id, message, type }]);
    setTimeout(() => {
      setNotifications(prev => prev.filter(n => n.id !== id));
    }, 5000);
  };

  const loadSurveys = async (retryCount = 0, maxRetries = 3, delay = 1000) => {
    if (!currentUser) return;
    setIsLoadingSurveys(true);
    const userId = currentUser.get_id();
    console.log('Loading surveys for userId:', userId); // Debug
    const filter = isSiteAdmin ? '' : `&$filter=Owners/Id eq ${userId} or Author/Id eq ${userId}`;
    try {
      const response = await $.ajax({
        url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items?$select=Id,Title,Owners/Id,Owners/Title,Author/Id,StartDate,EndDate,Status,Archive,surveyJson&$expand=Owners,Author${filter}`,
        headers: { "Accept": "application/json; odata=verbose" },
        xhrFields: { withCredentials: true }
      });
      console.log('Surveys API response (attempt ' + (retryCount + 1) + '):', response.d.results); // Debug
      const surveys = response.d.results.map(s => {
        let description = 'No description available';
        try {
          if (s.surveyJson) {
            const parsed = JSON.parse(s.surveyJson);
            description = parsed?.description || 'No description available';
          }
        } catch (e) {
          console.error(`Error parsing surveyJson for survey ${s.Id}:`, e);
        }
        return {
          ...s,
          Owners: { results: s.Owners ? s.Owners.results || [] : [] },
          Description: description
        };
      });
      const updatedSurveys = await Promise.all(surveys.map(s => 
        $.ajax({
          url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('SurveyResponses')/items?$filter=SurveyID eq ${s.Id}&$top=1&$inlinecount=allpages`,
          headers: { "Accept": "application/json; odata=verbose" },
          xhrFields: { withCredentials: true }
        }).then(res => ({ ...s, responseCount: res.d.__count || 0 }))
          .catch(error => {
            console.error(`Error fetching responses for survey ${s.Id}:`, error);
            addNotification(`Failed to load response count for survey "${s.Title}".`, 'error');
            return { ...s, responseCount: 0 };
          })
      ));
      console.log('Updated surveys:', updatedSurveys); // Debug
      setSurveys(updatedSurveys);
      setIsLoadingSurveys(false);
      if (updatedSurveys.length === 0) {
        addNotification(`No surveys found for user ID ${userId}. Ensure you are an owner or creator.`, 'warning');
      }
    } catch (error) {
      console.error('Error fetching surveys (attempt ' + (retryCount + 1) + '):', error);
      if (retryCount < maxRetries - 1) {
        console.log('Retrying loadSurveys in ' + delay + 'ms...');
        setTimeout(() => loadSurveys(retryCount + 1, maxRetries, delay * 2), delay);
      } else {
        addNotification(`Failed to load surveys after ${maxRetries} attempts. Ensure the "Surveys" list exists and you have read access.`, 'error');
        setIsLoadingSurveys(false);
      }
    }
  };

  useEffect(() => {
    setIsLoadingUser(true);
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', () => {
      const context = SP.ClientContext.get_current();
      const user = context.get_web().get_currentUser();
      context.load(user);
      context.executeQueryAsync(
        () => {
          setCurrentUser(user);
          $.ajax({
            url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/currentuser?$select=Id,IsSiteAdmin`,
            headers: { "Accept": "application/json; odata=verbose" },
            xhrFields: { withCredentials: true },
            success: (userData) => {
              const isSiteAdmin = userData.d.IsSiteAdmin;
              setIsSiteAdmin(isSiteAdmin);
              $.ajax({
                url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/currentuser/groups`,
                headers: { "Accept": "application/json; odata=verbose" },
                xhrFields: { withCredentials: true },
                success: (groupData) => {
                  const isOwnerGroup = groupData.d.results.some(g => g.Title.includes('Owners'));
                  setUserRole(isSiteAdmin || isOwnerGroup ? 'owner' : 'member');
                  setIsLoadingUser(false);
                  loadSurveys();
                },
                error: (xhr, status, error) => {
                  console.error('Error fetching groups:', error);
                  addNotification('Failed to load user groups.', 'error');
                  setIsLoadingUser(false);
                  loadSurveys();
                }
              });
            },
            error: (xhr, status, error) => {
              console.error('Error checking site admin status:', error);
              addNotification('Failed to check user permissions.', 'error');
              setIsLoadingUser(false);
              loadSurveys();
            }
          });
        },
        (sender, args) => {
          console.error('Error loading user:', args.get_message());
          addNotification('Failed to load user information.', 'error');
          setIsLoadingUser(false);
          loadSurveys();
        }
      );
    });
  }, []);

  const applyFilters = (survey) => {
    const { status, search } = filters;
    const matchesStatus = status.length === 0 || status.includes(survey.Status) || (survey.Archive && status.includes('Archive'));
    const matchesSearch = !search || survey.Title.toLowerCase().includes(search.toLowerCase());
    return matchesStatus && matchesSearch;
  };

  if (isLoadingUser) {
    return (
      <div className="flex items-center justify-center h-screen">
        <div className="animate-spin rounded-full h-12 w-12 border-t-4 border-blue-500"></div>
      </div>
    );
  }

  const filteredSurveys = surveys.filter(applyFilters);

  return (
    <div className="flex flex-col h-screen">
      <div className="fixed top-4 right-4 z-60 space-y-2">
        {notifications.map(n => (
          <Notification
            key={n.id}
            message={n.message}
            type={n.type}
            onClose={() => setNotifications(prev => prev.filter(notification => notification.id !== n.id))}
          />
        ))}
      </div>
      <TopNav username={currentUser?.get_title()} />
      <div className="flex flex-1">
        <SideNav 
          filters={filters} 
          onFilterChange={setFilters} 
          isOpen={isSideNavOpen} 
          toggle={() => setIsSideNavOpen(!isSideNavOpen)} 
          className={`lg:block ${isSideNavOpen ? 'block' : 'hidden'} md:w-1/4 bg-gray-100 p-4`} 
        />
        <div className="flex-1 p-4">
          <div className="mb-4">
            <button 
              onClick={() => window.open('builder.aspx', '_blank')} 
              className="bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600"
              aria-label="Create new survey form"
            >
              Create New Form
            </button>
          </div>
          {isLoadingSurveys ? (
            <div className="flex items-center justify-center h-full">
              <div className="animate-spin rounded-full h-12 w-12 border-t-4 border-blue-500 mr-4"></div>
              <span>Loading surveys...</span>
            </div>
          ) : filteredSurveys.length === 0 ? (
            <div className="flex items-center justify-center h-full">
              <span className="text-gray-500">No surveys available</span>
            </div>
          ) : (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {filteredSurveys.map(survey => (
                <SurveyCard 
                  key={survey.Id} 
                  survey={survey} 
                  userRole={userRole} 
                  currentUserId={currentUser?.get_id()} 
                  addNotification={addNotification} 
                  loadSurveys={loadSurveys}
                />
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

const Notification = ({ message, type, onClose }) => (
  <div className={`p-4 rounded shadow flex justify-between items-center ${type === 'success' ? 'bg-green-100 text-green-800' : type === 'warning' ? 'bg-yellow-100 text-yellow-800' : 'bg-red-100 text-red-800'}`}>
    <span>{message}</span>
    <button onClick={onClose} className="ml-4 text-lg font-bold" aria-label="Close notification">&times;</button>
  </div>
);

const TopNav = ({ username }) => (
  <nav className="bg-blue-600 p-4 flex justify-between items-center text-white">
    <div className="flex items-center">
      <img src="/SiteAssets/logo.png" alt="Logo" className="h-8" />
      <h1 className="ml-4">Survey Manager</h1>
    </div>
    <div className="text-right">{username}</div>
  </nav>
);

const SideNav = ({ filters, onFilterChange, isOpen, toggle, className }) => {
  const handleStatusChange = (e) => {
    const value = e.target.value;
    const newStatus = e.target.checked 
      ? [...filters.status, value]
      : filters.status.filter(s => s !== value);
    onFilterChange({ ...filters, status: newStatus });
  };

  return (
    <div className={className}>
      <button className="lg:hidden mb-4 p-2 bg-gray-200 rounded hover:bg-gray-300" onClick={toggle} aria-label={isOpen ? 'Close menu' : 'Open menu'}>
        {isOpen ? 'Close Menu' : 'Open Menu'}
      </button>
      <input 
        type="text" 
        placeholder="Search surveys..." 
        className="w-full p-2 mb-4 border rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
        onChange={(e) => onFilterChange({ ...filters, search: e.target.value })}
        aria-label="Search surveys"
      />
      <div className="space-y-2">
        <label className="flex items-center">
          <input type="checkbox" value="Publish" onChange={handleStatusChange} className="mr-2" aria-label="Filter by Published status" /> Published
        </label>
        <label className="flex items-center">
          <input type="checkbox" value="Draft" onChange={handleStatusChange} className="mr-2" aria-label="Filter by Draft status" /> Draft
        </label>
        <label className="flex items-center">
          <input type="checkbox" value="Archive" onChange={handleStatusChange} className="mr-2" aria-label="Filter by Archived status" /> Archived
        </label>
      </div>
    </div>
  );
};

const SurveyCard = ({ survey, userRole, currentUserId, addNotification, loadSurveys }) => {
  const [showQRModal, setShowQRModal] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const formUrl = `${_spPageContextInfo.webAbsoluteUrl}/SitePages/filler.aspx?surveyId=${survey.Id}`;

  const formatDate = (date) => date ? new Date(date).toLocaleDateString() : 'Not set';

  return (
    <div className="border p-4 rounded shadow bg-white hover:shadow-lg transition flex flex-col">
      <div className="flex-1">
        <h2 className="text-lg font-bold">{survey.Title}</h2>
        <p className="text-gray-600">{survey.Description}</p>
        <p>Responses: {survey.responseCount}</p>
        <p>Status: {survey.Status} {survey.Archive ? '(Archived)' : ''}</p>
        <p>Dates: {formatDate(survey.StartDate)} - {formatDate(survey.EndDate)}</p>
      </div>
      <div className="mt-4 flex flex-wrap gap-2 border-t pt-2">
        <button 
          className="flex items-center bg-blue-500 text-white px-3 py-1 rounded hover:bg-blue-600" 
          onClick={() => window.open(`builder.aspx?surveyId=${survey.Id}`, '_blank')}
          title="Edit the survey form"
          aria-label="Edit survey form"
        >
          <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
          </svg>
          Edit Form
        </button>
        <button 
          className="flex items-center bg-yellow-500 text-white px-3 py-1 rounded hover:bg-yellow-600" 
          onClick={() => window.open(`report.aspx?surveyId=${survey.Id}`, '_blank')}
          title="View survey report"
          aria-label="View survey report"
        >
          <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 17v-2m0-2v-2m0-2V7m6 10v-2m0-2v-2m0-2V7m-6-2h6m4 0H5a2 2 0 00-2 2v12a2 2 0 002 2h14a2 2 0 002-2V7a2 2 0 00-2-2z"></path>
          </svg>
          View Report
        </button>
        <button 
          className="flex items-center bg-purple-500 text-white px-3 py-1 rounded hover:bg-purple-600" 
          onClick={() => setShowQRModal(true)}
          title="Generate QR code"
          aria-label="Generate QR code"
        >
          <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 4v1m6 11h2m-6 0h-2v4m0-11v3m-2 4h2M6 12H4m2 4v4m0-11v3m-2 4h2m7-7h3m-3 3h3m-3 3h3"></path>
          </svg>
          QR Code
        </button>
        <button 
          className="flex items-center bg-gray-500 text-white px-3 py-1 rounded hover:bg-gray-600" 
          onClick={() => setShowEditModal(true)}
          title="Edit survey metadata"
          aria-label="Edit survey metadata"
        >
          <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.572L16.732 3.732z"></path>
          </svg>
          Edit Metadata
        </button>
        <button 
          className="flex items-center bg-green-500 text-white px-3 py-1 rounded hover:bg-green-600" 
          onClick={() => window.open(`filler.aspx?surveyId=${survey.Id}`, '_blank')}
          title="Fill out the survey"
          aria-label="Fill out survey"
        >
          <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4M7.835 4.697a3.5 3.5 0 105.33 4.606 3.5 3.5 0 01-5.33-4.606zM12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707"></path>
          </svg>
          Fill Form
        </button>
      </div>
      {showQRModal && <QRModal url={formUrl} onClose={() => setShowQRModal(false)} addNotification={addNotification} />}
      {showEditModal && <EditModal survey={survey} onClose={() => setShowEditModal(false)} addNotification={addNotification} currentUserId={currentUserId} />}
    </div>
  );
};

const QRModal = ({ url, onClose, addNotification }) => {
  const qrRef = React.useRef(null);
  useEffect(() => {
    new QRious({ element: qrRef.current, value: url, size: 200 });
  }, [url]);

  const downloadQR = () => {
    const link = document.createElement('a');
    link.href = qrRef.current.toDataURL();
    link.download = 'qrcode.png';
    link.click();
  };

  return (
    <div className="fixed inset-0 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-md">
        <div className="flex justify-between items-center p-4 border-b">
          <h2 className="text-lg font-bold">QR Code</h2>
          <button onClick={onClose} className="text-gray-600 hover:text-gray-800" aria-label="Close QR code modal">
            &times;
          </button>
        </div>
        <div className="p-6">
          <canvas ref={qrRef} className="mx-auto"></canvas>
        </div>
        <div className="flex gap-2 justify-end p-4 border-t">
          <button 
            className="bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600" 
            onClick={downloadQR}
            aria-label="Download QR code"
          >
            Download
          </button>
          <button 
            className="bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600" 
            onClick={() => navigator.clipboard.writeText(url).then(() => addNotification('URL copied to clipboard!'))}
            aria-label="Copy QR code URL"
          >
            Copy URL
          </button>
          <button 
            className="bg-red-500 text-white px-4 py-2 rounded hover:bg-red-600" 
            onClick={onClose}
            aria-label="Close modal"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
};

const EditModal = ({ survey, onClose, addNotification, currentUserId }) => {
  const [form, setForm] = useState({
    Owners: Array.isArray(survey.Owners?.results) ? survey.Owners.results.map(o => ({ Id: o.Id, Title: o.Title })) : [],
    StartDate: survey.StartDate ? new Date(survey.StartDate).toISOString().split('T')[0] : '',
    EndDate: survey.EndDate ? new Date(survey.EndDate).toISOString().split('T')[0] : '',
    Status: survey.Status || 'Draft',
    Archive: survey.Archive || false
  });
  const [searchTerm, setSearchTerm] = useState('');
  const [searchResults, setSearchResults] = useState([]);
  const [isLoadingUsers, setIsLoadingUsers] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [showDropdown, setShowDropdown] = useState(false);

  useEffect(() => {
    if (!searchTerm) {
      setSearchResults([]);
      setShowDropdown(false);
      return;
    }

    const debounce = setTimeout(() => {
      setIsLoadingUsers(true);
      $.ajax({
        url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/siteusers?$select=Id,Title&$filter=substringof('${encodeURIComponent(searchTerm)}',Title)&$top=10`,
        headers: { "Accept": "application/json; odata=verbose" },
        xhrFields: { withCredentials: true },
        success: (data) => {
          const users = data.d.results
            .filter(u => u.Id && u.Title)
            .map(u => ({ Id: u.Id, Title: u.Title }));
          const availableUsers = users.filter(u => !form.Owners.some(selected => selected.Id === u.Id));
          setSearchResults(availableUsers);
          setIsLoadingUsers(false);
          setShowDropdown(true);
        },
        error: (xhr, status, error) => {
          console.error('Error searching users:', error);
          addNotification('Failed to search users.', 'error');
          setIsLoadingUsers(false);
          setShowDropdown(false);
        }
      });
    }, 300);

    return () => clearTimeout(debounce);
  }, [searchTerm, form.Owners]);

  const handleUserSelect = (user) => {
    setForm(prev => ({ ...prev, Owners: [...prev.Owners, user] }));
    setSearchTerm('');
    setShowDropdown(false);
  };

  const handleUserRemove = (userId) => {
    if (userId === currentUserId) {
      addNotification('You cannot remove yourself as an owner.', 'error');
      return;
    }
    setForm(prev => ({ ...prev, Owners: prev.Owners.filter(o => o.Id !== userId) }));
  };

  const handleSave = async () => {
    setIsSaving(true);
    try {
      if (!form.Owners.some(o => o.Id === currentUserId)) {
        throw new Error('You must remain an owner of the survey.');
      }

      const digest = await getDigest();

      const payload = {
        '__metadata': { 'type': 'SP.Data.SurveysListItem' },
        OwnersId: { results: form.Owners.map(o => o.Id) },
        Status: form.Status,
        Archive: form.Archive
      };
      if (form.StartDate) {
        payload.StartDate = new Date(form.StartDate).toISOString();
      }
      if (form.EndDate) {
        payload.EndDate = new Date(form.EndDate).toISOString();
      }
      console.log('Saving metadata for survey:', survey.Id, payload); // Debug

      await $.ajax({
        url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items(${survey.Id})`,
        type: 'POST',
        data: JSON.stringify(payload),
        headers: {
          "X-HTTP-Method": "MERGE",
          "If-Match": "*",
          "Accept": "application/json; odata=verbose",
          "Content-Type": "application/json; odata=verbose",
          "X-RequestDigest": digest
        },
        xhrFields: { withCredentials: true }
      });

      const permissions = await $.ajax({
        url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items(${survey.Id})/effectiveBasePermissions`,
        headers: { "Accept": "application/json; odata=verbose" },
        xhrFields: { withCredentials: true }
      });
      const hasManagePermissions = permissions.d.EffectiveBasePermissions.High & 0x00000080;

      if (hasManagePermissions) {
        await $.ajax({
          url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items(${survey.Id})/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
          type: 'POST',
          headers: {
            "Accept": "application/json; odata=verbose",
            "X-RequestDigest": digest
          },
          xhrFields: { withCredentials: true }
        });

        if (form.Owners.length > 0) {
          await Promise.all(form.Owners.map(user => 
            $.ajax({
              url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items(${survey.Id})/roleassignments/addroleassignment(principalid=${user.Id}, roledefid=1073741827)`,
              type: 'POST',
              headers: {
                "Accept": "application/json; odata=verbose",
                "X-RequestDigest": digest
              },
              xhrFields: { withCredentials: true }
            })
          ));
        }
        addNotification('Survey metadata and permissions updated successfully!');
      } else {
        addNotification('Survey metadata updated. Permissions not modified due to insufficient access.', 'warning');
      }

      console.log('Metadata save successful for survey:', survey.Id); // Debug
      setTimeout(() => location.reload(), 500);
      onClose();
    } catch (error) {
      console.error('Error updating survey:', error);
      let errorMessage = error.message || error.responseText || 'Unknown error';
      if (error.status === 403) {
        errorMessage = 'Access denied. Ensure you have Manage Permissions on this survey.';
      } else if (errorMessage.includes('Invalid Form Digest')) {
        errorMessage = 'Invalid or expired request digest token. Please try again.';
      }
      addNotification(`Failed to update survey: ${errorMessage}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-md">
        <div className="flex justify-between items-center p-4 border-b">
          <h2 className="text-lg font-bold">Edit Metadata</h2>
          <button onClick={onClose} className="text-gray-600 hover:text-gray-800" aria-label="Close metadata modal">
            &times;
          </button>
        </div>
        <div className="p-6 max-h-96 overflow-y-auto">
          <div className="space-y-4">
            <div>
              <label className="block mb-1">Owners</label>
              <div className="relative">
                <input
                  type="text"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  placeholder="Search for users by name..."
                  className="w-full p-2 border rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
                  aria-label="Search for users"
                />
                {isLoadingUsers && (
                  <div className="absolute top-2 right-2">
                    <div className="animate-spin rounded-full h-5 w-5 border-t-2 border-blue-500"></div>
                  </div>
                )}
                {showDropdown && searchResults.length > 0 && (
                  <ul className="absolute z-10 w-full bg-white border rounded mt-1 max-h-48 overflow-y-auto shadow-lg">
                    {searchResults.map(user => (
                      <li
                        key={user.Id}
                        onClick={() => handleUserSelect(user)}
                        className="p-2 hover:bg-gray-100 cursor-pointer border-b last:border-b-0"
                        role="option"
                        aria-selected="false"
                      >
                        {user.Title}
                      </li>
                    ))}
                  </ul>
                )}
              </div>
              <div className="mt-2 flex flex-wrap gap-2">
                {form.Owners.length === 0 ? (
                  <p className="text-gray-500 text-sm">No owners selected</p>
                ) : (
                  form.Owners.map(user => (
                    <div
                      key={user.Id}
                      className="flex items-center bg-blue-100 text-blue-800 px-2 py-1 rounded-full text-sm"
                    >
                      <span>{user.Title}</span>
                      <button
                        onClick={() => handleUserRemove(user.Id)}
                        className="ml-2 text-red-600 hover:text-red-800 font-bold"
                        disabled={user.Id === currentUserId}
                        aria-label={`Remove ${user.Title} from owners`}
                      >
                        {user.Id === currentUserId ? '' : '&times;'}
                      </button>
                    </div>
                  ))
                )}
              </div>
            </div>
            <div>
              <label className="block mb-1">Start Date</label>
              <input
                type="date"
                value={form.StartDate}
                onChange={(e) => setForm(prev => ({ ...prev, StartDate: e.target.value }))}
                className="w-full p-2 border rounded"
                aria-label="Start date"
              />
            </div>
            <div>
              <label className="block mb-1">End Date</label>
              <input
                type="date"
                value={form.EndDate}
                onChange={(e) => setForm(prev => ({ ...prev, EndDate: e.target.value }))}
                className="w-full p-2 border rounded"
                aria-label="End date"
              />
            </div>
            <div>
              <label className="block mb-1">Status</label>
              <select
                value={form.Status}
                onChange={(e) => setForm(prev => ({ ...prev, Status: e.target.value }))}
                className="w-full p-2 border rounded"
                aria-label="Survey status"
              >
                <option value="Publish">Publish</option>
                <option value="Draft">Draft</option>
              </select>
            </div>
            <div>
              <label className="flex items-center">
                <input
                  type="checkbox"
                  checked={form.Archive}
                  onChange={(e) => setForm(prev => ({ ...prev, Archive: e.target.checked }))}
                  className="mr-2"
                  aria-label="Archive survey"
                />
                Archive
              </label>
            </div>
          </div>
        </div>
        <div className="flex gap-2 justify-end p-4 border-t">
          <button 
            className={`bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600 flex items-center ${isSaving ? 'opacity-50 cursor-not-allowed' : ''}`} 
            onClick={handleSave}
            disabled={isSaving}
            aria-label="Save metadata"
          >
            {isSaving ? (
              <>
                <div className="animate-spin rounded-full h-5 w-5 border-t-2 border-white mr-2"></div>
                Saving...
              </>
            ) : (
              'Save'
            )}
          </button>
          <button 
            className="bg-red-500 text-white px-4 py-2 rounded hover:bg-red-600" 
            onClick={onClose}
            disabled={isSaving}
            aria-label="Cancel metadata edit"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
};

ReactDOM.render(<App />, document.getElementById('root'));