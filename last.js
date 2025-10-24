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

  const loadSurveys = () => {
    setIsLoadingSurveys(true);
    const filter = currentUser ? `&$filter=Owners/Id eq ${currentUser.get_id()} and Author/Id eq ${currentUser.get_id()}` : '';
    $.ajax({
      url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items?$select=Id,Title,Owners/Id,Owners/Title,Author/Id,StartDate,EndDate,Status,Archive&$expand=Owners,Author${filter}`,
      headers: { "Accept": "application/json; odata=verbose" },
      xhrFields: { withCredentials: true },
      success: (data) => {
        console.log('Surveys API response:', data.d.results); // Debug: Log API response
        const surveys = data.d.results.map(s => ({
          ...s,
          Owners: { results: s.Owners ? s.Owners.results || [] : [] }
        }));
        Promise.all(surveys.map(s => 
          $.ajax({
            url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('SurveyResponses')/items?$filter=SurveyID eq ${s.Id}&$top=1&$inlinecount=allpages`,
            headers: { "Accept": "application/json; odata=verbose" },
            xhrFields: { withCredentials: true },
          }).then(res => ({ ...s, responseCount: res.d.__count || 0 }))
            .catch(error => {
              console.error(`Error fetching responses for survey ${s.Id}:`, error);
              addNotification(`Failed to load response count for survey "${s.Title}".`, 'error');
              return { ...s, responseCount: 0 };
            })
        )).then(updatedSurveys => {
          console.log('Updated surveys:', updatedSurveys); // Debug: Log final surveys
          setSurveys(updatedSurveys);
          setIsLoadingSurveys(false);
          if (updatedSurveys.length === 0) {
            addNotification('No surveys found where you are the creator and owner.', 'warning');
          }
        });
      },
      error: (xhr, status, error) => {
        console.error('Error fetching surveys:', error);
        addNotification('Failed to load surveys. Ensure the "Surveys" list exists.', 'error');
        setIsLoadingSurveys(false);
      }
    });
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
      <div className="fixed top-4 right-4 z-50 space-y-2">
        {notifications.map(n => (
          <Notification
            key={n.id}
            message={n.message}
            type={n.type}
            onClose={() => setNotifications(prev => prev.filter(notification => notification.id !== n.id))}
          />
        ))}
      </div>
      <TopNav username={currentUser?.get_title()} onCreate={() => window.open('builder.aspx', '_blank')} />
      <div className="flex flex-1">
        <SideNav 
          filters={filters} 
          onFilterChange={setFilters} 
          isOpen={isSideNavOpen} 
          toggle={() => setIsSideNavOpen(!isSideNavOpen)} 
          className={`lg:block ${isSideNavOpen ? 'block' : 'hidden'} md:w-1/4 bg-gray-100 p-4`} 
        />
        <div className="flex-1 p-4">
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
    <button onClick={onClose} className="ml-4 text-lg font-bold">&times;</button>
  </div>
);

const TopNav = ({ username, onCreate }) => (
  <nav className="bg-blue-600 p-4 flex justify-between items-center text-white">
    <div className="flex items-center">
      <img src="/SiteAssets/logo.png" alt="Logo" className="h-8" />
      <h1 className="ml-4">Survey Manager</h1>
    </div>
    <div>{username}</div>
    <button onClick={onCreate} className="bg-green-500 px-4 py-2 rounded hover:bg-green-600">Create New Form</button>
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
      <button className="lg:hidden mb-4 p-2 bg-gray-200 rounded hover:bg-gray-300" onClick={toggle}>
        {isOpen ? 'Close Menu' : 'Open Menu'}
      </button>
      <input 
        type="text" 
        placeholder="Search surveys..." 
        className="w-full p-2 mb-4 border rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
        onChange={(e) => onFilterChange({ ...filters, search: e.target.value })}
      />
      <div className="space-y-2">
        <label className="flex items-center">
          <input type="checkbox" value="Publish" onChange={handleStatusChange} className="mr-2" /> Published
        </label>
        <label className="flex items-center">
          <input type="checkbox" value="Draft" onChange={handleStatusChange} className="mr-2" /> Draft
        </label>
        <label className="flex items-center">
          <input type="checkbox" value="Archive" onChange={handleStatusChange} className="mr-2" /> Archived
        </label>
      </div>
    </div>
  );
};

const SurveyCard = ({ survey, userRole, currentUserId, addNotification, loadSurveys }) => {
  const [showQRModal, setShowQRModal] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const formUrl = `${_spPageContextInfo.webAbsoluteUrl}/SitePages/filler.aspx?surveyId=${survey.Id}`;

  return (
    <div className="border p-4 rounded shadow bg-white hover:shadow-lg transition">
      <h2 className="text-lg font-bold">{survey.Title}</h2>
      <p>Responses: {survey.responseCount}</p>
      <p>Status: {survey.Status} {survey.Archive ? '(Archived)' : ''}</p>
      <div className="flex flex-wrap gap-2 mt-2">
        <button 
          className="bg-blue-500 text-white px-3 py-1 rounded hover:bg-blue-600" 
          onClick={() => window.open(`builder.aspx?surveyId=${survey.Id}`, '_blank')}
        >
          Edit Form
        </button>
        <button 
          className="bg-yellow-500 text-white px-3 py-1 rounded hover:bg-yellow-600" 
          onClick={() => window.open(`report.aspx?surveyId=${survey.Id}`, '_blank')}
        >
          View Report
        </button>
        <button 
          className="bg-purple-500 text-white px-3 py-1 rounded hover:bg-purple-600" 
          onClick={() => setShowQRModal(true)}
        >
          QR Code
        </button>
        <button 
          className="bg-gray-500 text-white px-3 py-1 rounded hover:bg-gray-600" 
          onClick={() => setShowEditModal(true)}
        >
          Edit Metadata
        </button>
        <button 
          className="bg-green-500 text-white px-3 py-1 rounded hover:bg-green-600" 
          onClick={() => window.open(`filler.aspx?surveyId=${survey.Id}`, '_blank')}
        >
          Fill Form
        </button>
      </div>
      {showQRModal && <QRModal url={formUrl} onClose={() => setShowQRModal(false)} addNotification={addNotification} />}
      {showEditModal && <EditModal survey={survey} onClose={() => setShowEditModal(false)} onSave={() => loadSurveys()} addNotification={addNotification} currentUserId={currentUserId} />}
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
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white p-6 rounded-lg shadow-xl">
        <canvas ref={qrRef} className="mx-auto"></canvas>
        <div className="mt-4 flex gap-2 justify-center">
          <button 
            className="bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600" 
            onClick={downloadQR}
          >
            Download
          </button>
          <button 
            className="bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600" 
            onClick={() => navigator.clipboard.writeText(url).then(() => addNotification('URL copied to clipboard!'))}
          >
            Copy URL
          </button>
          <button 
            className="bg-red-500 text-white px-4 py-2 rounded hover:bg-red-600" 
            onClick={onClose}
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
};

const EditModal = ({ survey, onClose, onSave, addNotification, currentUserId }) => {
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
      // Ensure current user remains in Owners
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
      console.log('Saving payload:', payload);

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

      if (typeof onSave === 'function') {
        onSave();
      } else {
        console.warn('onSave is not a function; skipping survey refresh');
      }
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
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white p-6 rounded-lg shadow-xl w-full max-w-md max-h-96 overflow-y-auto">
        <h2 className="text-lg font-bold mb-4">Edit Metadata</h2>
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
            />
          </div>
          <div>
            <label className="block mb-1">End Date</label>
            <input
              type="date"
              value={form.EndDate}
              onChange={(e) => setForm(prev => ({ ...prev, EndDate: e.target.value }))}
              className="w-full p-2 border rounded"
            />
          </div>
          <div>
            <label className="block mb-1">Status</label>
            <select
              value={form.Status}
              onChange={(e) => setForm(prev => ({ ...prev, Status: e.target.value }))}
              className="w-full p-2 border rounded"
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
              />
              Archive
            </label>
          </div>
        </div>
        <div className="mt-6 flex gap-2 justify-end">
          <button 
            className={`bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600 flex items-center ${isSaving ? 'opacity-50 cursor-not-allowed' : ''}`} 
            onClick={handleSave}
            disabled={isSaving}
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
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
};

ReactDOM.render(<App />, document.getElementById('root'));