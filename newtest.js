// app.js
const { useState, useEffect } = React;

const App = () => {
  const [surveys, setSurveys] = useState([]);
  const [userRole, setUserRole] = useState('');
  const [currentUser, setCurrentUser] = useState(null);
  const [filters, setFilters] = useState({ status: [], search: '' });
  const [isSideNavOpen, setIsSideNavOpen] = useState(false);
  const [isLoadingUser, setIsLoadingUser] = useState(true);
  const [isLoadingSurveys, setIsLoadingSurveys] = useState(false);
  const [notifications, setNotifications] = useState([]); // [{ id, message, type }]

  const addNotification = (message, type = 'success') => {
    const id = Date.now();
    setNotifications(prev => [...prev, { id, message, type }]);
    setTimeout(() => {
      setNotifications(prev => prev.filter(n => n.id !== id));
    }, 5000);
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
            url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/currentuser/groups`,
            headers: { "Accept": "application/json; odata=verbose" },
            success: (data) => {
              const isOwner = data.d.results.some(g => g.Title.includes('Owners'));
              setUserRole(isOwner ? 'owner' : 'member');
              setIsLoadingUser(false);
              loadSurveys(isOwner ? 'all' : 'owned');
            },
            error: (xhr, status, error) => {
              console.error('Error fetching groups:', error);
              addNotification('Failed to load user information.', 'error');
              setIsLoadingUser(false);
            }
          });
        },
        (sender, args) => {
          console.error('Error loading user:', args.get_message());
          addNotification('Failed to load user information.', 'error');
          setIsLoadingUser(false);
        }
      );
    });
  }, []);

  const loadSurveys = (mode) => {
    setIsLoadingSurveys(true);
    const filter = mode === 'owned' ? `&$filter=Owners/Id eq ${currentUser.get_id()}` : '';
    $.ajax({
      url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items?$select=Id,Title,Owners/Title,StartDate,EndDate,Status,Archive&$expand=Owners${filter}`,
      headers: { "Accept": "application/json; odata=verbose" },
      success: (data) => {
        const surveys = data.d.results;
        Promise.all(surveys.map(s => 
          $.ajax({
            url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Responses')/items?$filter=SurveyID eq ${s.Id}&$top=1&$inlinecount=allpages`,
            headers: { "Accept": "application/json; odata=verbose" }
          }).then(res => ({ ...s, responseCount: res.d.__count || 0 }))
        )).then(updatedSurveys => {
          setSurveys(updatedSurveys);
          setIsLoadingSurveys(false);
        });
      },
      error: (xhr, status, error) => {
        console.error('Error fetching surveys:', error);
        addNotification('Failed to load surveys.', 'error');
        setIsLoadingSurveys(false);
      }
    });
  };

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
          ) : (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {surveys.filter(applyFilters).map(survey => (
                <SurveyCard key={survey.Id} survey={survey} userRole={userRole} currentUserId={currentUser?.get_id()} addNotification={addNotification} />
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

const Notification = ({ message, type, onClose }) => (
  <div className={`p-4 rounded shadow flex justify-between items-center ${type === 'success' ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>
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

const SurveyCard = ({ survey, userRole, currentUserId, addNotification }) => {
  const [showQRModal, setShowQRModal] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const formUrl = `${_spPageContextInfo.webAbsoluteUrl}/SitePages/filler.aspx?surveyId=${survey.Id}`;

  const isOwner = userRole === 'owner' || survey.Owners.results.some(o => o.Id === currentUserId);

  return (
    <div className="border p-4 rounded shadow bg-white hover:shadow-lg transition">
      <h2 className="text-lg font-bold">{survey.Title}</h2>
      <p>Responses: {survey.responseCount}</p>
      <p>Status: {survey.Status} {survey.Archive ? '(Archived)' : ''}</p>
      <div className="flex flex-wrap gap-2 mt-2">
        {isOwner && (
          <>
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
          </>
        )}
        <button 
          className="bg-green-500 text-white px-3 py-1 rounded hover:bg-green-600" 
          onClick={() => window.open(`filler.aspx?surveyId=${survey.Id}`, '_blank')}
        >
          Fill Form
        </button>
      </div>
      {showQRModal && <QRModal url={formUrl} onClose={() => setShowQRModal(false)} addNotification={addNotification} />}
      {showEditModal && <EditModal survey={survey} onClose={() => setShowEditModal(false)} onSave={() => loadSurveys(userRole)} addNotification={addNotification} />}
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

const EditModal = ({ survey, onClose, onSave, addNotification }) => {
  const [form, setForm] = useState({
    Owners: survey.Owners.results.map(o => o.Id),
    StartDate: survey.StartDate ? new Date(survey.StartDate).toISOString().split('T')[0] : '',
    EndDate: survey.EndDate ? new Date(survey.EndDate).toISOString().split('T')[0] : '',
    Status: survey.Status,
    Archive: survey.Archive
  });
  const [users, setUsers] = useState([]);
  const [isLoadingUsers, setIsLoadingUsers] = useState(true);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    setIsLoadingUsers(true);
    $.ajax({
      url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/siteusers?$select=Id,Title`,
      headers: { "Accept": "application/json; odata=verbose" },
      success: (data) => {
        setUsers(data.d.results.filter(u => u.Title));
        setIsLoadingUsers(false);
      },
      error: (xhr, status, error) => {
        console.error('Error fetching users:', error);
        addNotification('Failed to load users.', 'error');
        setIsLoadingUsers(false);
      }
    });
  }, []);

  const handleSave = () => {
    setIsSaving(true);
    $.ajax({
      url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items(${survey.Id})`,
      type: 'POST',
      data: JSON.stringify({
        '__metadata': { 'type': 'SP.Data.SurveysListItem' },
        OwnersId: { results: form.Owners },
        StartDate: form.StartDate || null,
        EndDate: form.EndDate || null,
        Status: form.Status,
        Archive: form.Archive
      }),
      headers: {
        "X-HTTP-Method": "MERGE",
        "If-Match": "*",
        "Accept": "application/json; odata=verbose",
        "Content-Type": "application/json; odata=verbose"
      },
      success: () => {
        $.ajax({
          url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items(${survey.Id}/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
          type: 'POST',
          headers: { "Accept": "application/json; odata=verbose" },
          success: () => {
            Promise.all(form.Owners.map(userId => 
              $.ajax({
                url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('Surveys')/items(${survey.Id}/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741827)`,
                type: 'POST',
                headers: { "Accept": "application/json; odata=verbose" }
              })
            )).then(() => {
              setIsSaving(false);
              addNotification('Survey metadata updated successfully!');
              onSave();
              onClose();
            }).catch(error => {
              console.error('Error setting permissions:', error);
              addNotification('Failed to update survey metadata.', 'error');
              setIsSaving(false);
            });
          },
          error: (xhr, status, error) => {
            console.error('Error setting permissions:', error);
            addNotification('Failed to update survey metadata.', 'error');
            setIsSaving(false);
          }
        });
      },
      error: (xhr, status, error) => {
        console.error('Error updating survey:', error);
        addNotification('Failed to update survey metadata.', 'error');
        setIsSaving(false);
      }
    });
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white p-6 rounded-lg shadow-xl w-full max-w-md">
        <h2 className="text-lg font-bold mb-4">Edit Metadata</h2>
        <div className="space-y-4">
          <div>
            <label className="block mb-1">Owners</label>
            {isLoadingUsers ? (
              <div className="flex items-center">
                <div className="animate-spin rounded-full h-6 w-6 border-t-2 border-blue-500 mr-2"></div>
                <span>Loading users...</span>
              </div>
            ) : (
              <select
                multiple
                value={form.Owners}
                onChange={(e) => setForm({ ...form, Owners: Array.from(e.target.selectedOptions, o => parseInt(o.value)) })}
                className="w-full p-2 border rounded"
              >
                {users.map(user => (
                  <option key={user.Id} value={user.Id}>{user.Title}</option>
                ))}
              </select>
            )}
          </div>
          <div>
            <label className="block mb-1">Start Date</label>
            <input
              type="date"
              value={form.StartDate}
              onChange={(e) => setForm({ ...form, StartDate: e.target.value })}
              className="w-full p-2 border rounded"
            />
          </div>
          <div>
            <label className="block mb-1">End Date</label>
            <input
              type="date"
              value={form.EndDate}
              onChange={(e) => setForm({ ...form, EndDate: e.target.value })}
              className="w-full p-2 border rounded"
            />
          </div>
          <div>
            <label className="block mb-1">Status</label>
            <select
              value={form.Status}
              onChange={(e) => setForm({ ...form, Status: e.target.value })}
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
                onChange={(e) => setForm({ ...form, Archive: e.target.checked })}
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