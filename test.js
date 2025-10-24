$(document).ready(function() {
    try {
        let attempts = 0;
        const maxAttempts = 50; // 5 seconds timeout
        const checkSP = setInterval(() => {
            if (typeof SP !== 'undefined' && SP.SOD) {
                clearInterval(checkSP);
                SP.SOD.executeOrDelayUntilScriptLoaded(initApp, 'sp.js');
            } else if (attempts >= maxAttempts) {
                clearInterval(checkSP);
                console.error('SP.js failed to load after timeout.');
                alert('SharePoint scripts failed to load. Refresh or check network.');
            }
            attempts++;
        }, 100);
    } catch (e) {
        console.error('Page init failed:', e);
        alert('Failed to initialize dashboard.');
    }
});

function initApp() {
    console.log('initApp called');
    if (!SP.ClientContext) {
        console.error('SP.ClientContext not available');
        alert('SharePoint context failed to load. Check script references or refresh.');
        return;
    }

    if (!React || !ReactDOM) {
        console.error('React or ReactDOM not loaded');
        alert('React scripts failed to load. Refresh or check CDNs.');
        return;
    }

    if (!window.QRious) {
        console.error('QRious not loaded');
        alert('QRious script failed to load. Refresh or check CDN.');
        return;
    }

    console.log('Rendering React app');
    const App = () => {
        const [isManager, setIsManager] = React.useState(false);
        const [surveys, setSurveys] = React.useState([]);
        const [filteredSurveys, setFilteredSurveys] = React.useState([]);
        const [statusFilters, setStatusFilters] = React.useState({
            All: true,
            draft: false,
            published: false,
            Upcoming: false,
            Active: false,
            Expired: false,
            Archived: false
        });
        const [searchQuery, setSearchQuery] = React.useState('');
        const [totalSurveys, setTotalSurveys] = React.useState(0);
        const [totalResponses, setTotalResponses] = React.useState(0);
        const [selectedId, setSelectedId] = React.useState(null); // Used for duplicate
        const [showDuplicateDialog, setShowDuplicateDialog] = React.useState(false); // Removed showDeleteDialog
        const [showShareDialog, setShowShareDialog] = React.useState(false);
        const [showEditMetadataDialog, setShowEditMetadataDialog] = React.useState(false);
        const [selectedSurvey, setSelectedSurvey] = React.useState(null);
        const [users, setUsers] = React.useState([]);
        const [shareLink, setShareLink] = React.useState('');
        const [username, setUsername] = React.useState('Loading...');
        const qrContainerRef = React.useRef(null);

        React.useEffect(() => {
            loadUserInfo();
            loadUserRole();
            loadSiteUsers();
            loadSurveysData();
            fetchCounts();
        }, []);

        function loadUserInfo() {
            try {
                const context = SP.ClientContext.get_current();
                const user = context.get_web().get_currentUser();
                context.load(user);
                context.executeQueryAsync(
                    () => setUsername(user.get_title()),
                    onFail('Failed to load user info')
                );
            } catch (e) {
                console.error('User info error:', e);
                setUsername('Error');
            }
        }

        function loadUserRole() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const user = web.get_currentUser();
                const groups = user.get_groups();
                context.load(groups);
                context.executeQueryAsync(
                    () => {
                        let isAdmin = false;
                        const enumerator = groups.getEnumerator();
                        while (enumerator.moveNext()) {
                            if (enumerator.get_current().get_title().includes('Owners')) {
                                isAdmin = true;
                                break;
                            }
                        }
                        setIsManager(isAdmin);
                    },
                    onFail('Failed to load user role')
                );
            } catch (e) {
                console.error('User role error:', e);
                setIsManager(false);
            }
        }

        function loadSiteUsers() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const users = web.get_siteUsers();
                context.load(users);
                context.executeQueryAsync(
                    () => {
                        const userList = [];
                        const enumerator = users.getEnumerator();
                        while (enumerator.moveNext()) {
                            const user = enumerator.get_current();
                            if (user.get_userId()) {
                                userList.push({
                                    id: user.get_id(),
                                    title: user.get_title(),
                                    login: user.get_loginName()
                                });
                            }
                        }
                        setUsers(userList);
                    },
                    onFail('Failed to load site users')
                );
            } catch (e) {
                console.error('Site users error:', e);
                setUsers([]);
            }
        }

        function loadSurveysData() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const currentUser = web.get_currentUser();
                const surveyList = web.get_lists().getByTitle('Surveys');
                context.load(currentUser);
                context.executeQueryAsync(
                    () => {
                        const query = new SP.CamlQuery();
                        query.set_viewXml(`
                            <View>
                                <Query>
                                    <Where>
                                        <Contains>
                                            <FieldRef Name="Owner"/>
                                            <Value Type="UserMulti">${currentUser.get_title()}</Value>
                                        </Contains>
                                    </Where>
                                </Query>
                            </View>
                        `);
                        const items = surveyList.getItems(query);
                        context.load(items, 'Include(Id, Title, Owner, Status, StartDate, EndDate, Archived, surveyData)');
                        context.executeQueryAsync(
                            () => {
                                console.log('Surveys loaded successfully'); // Governance: Log load
                                const surveyData = [];
                                const enumerator = items.getEnumerator();
                                while (enumerator.moveNext()) {
                                    const item = enumerator.get_current();
                                    const owners = item.get_item('Owner');
                                    const ownerNames = owners ? owners.map(o => o.get_lookupValue()).join(', ') : 'Unknown';
                                    surveyData.push({
                                        id: item.get_id(),
                                        title: item.get_item('Title') || 'Untitled',
                                        owner: ownerNames,
                                        created: item.get_item('Created') ? new Date(item.get_item('Created')).toLocaleDateString() : 'Unknown',
                                        status: item.get_item('Status') || 'draft',
                                        startDate: item.get_item('StartDate') ? new Date(item.get_item('StartDate')).toLocaleDateString() : 'No Start Date',
                                        endDate: item.get_item('EndDate') ? new Date(item.get_item('EndDate')).toLocaleDateString() : 'No End Date',
                                        archived: item.get_item('Archived') || false,
                                        responses: 0 // Updated below
                                    });
                                }
                                setSurveys(surveyData);
                                setFilteredSurveys(surveyData);
                            },
                            onFail('Failed to load surveys')
                        );
                    },
                    onFail('Failed to load current user')
                );
            } catch (e) {
                console.error('Surveys load error:', e);
                setSurveys([]);
                setFilteredSurveys([]);
                alert('Error loading surveys.');
            }
        }

        function fetchCounts() {
            try {
                const context = SP.ClientContext.get_current();
                const responseList = context.get_web().get_lists().getByTitle('SurveyResponses');
                const query = new SP.CamlQuery();
                query.set_viewXml('<View><Query></Query></View>');
                const items = responseList.getItems(query);
                context.load(items, 'Include(SurveyID)');
                context.executeQueryAsync(
                    () => {
                        console.log('Responses counted successfully'); // Governance: Log count
                        const responseCounts = {};
                        const enumerator = items.getEnumerator();
                        while (enumerator.moveNext()) {
                            const item = enumerator.get_current();
                            const surveyId = item.get_item('SurveyID');
                            responseCounts[surveyId] = (responseCounts[surveyId] || 0) + 1;
                        }
                        const updatedSurveys = surveys.map(s => ({
                            ...s,
                            responses: responseCounts[s.id] || 0
                        }));
                        setSurveys(updatedSurveys);
                        setFilteredSurveys(updatedSurveys);
                        setTotalSurveys(updatedSurveys.length);
                        setTotalResponses(Object.values(responseCounts).reduce((a, b) => a + b, 0));
                    },
                    onFail('Failed to fetch response counts')
                );
            } catch (e) {
                console.error('Response counts error:', e);
                setTotalSurveys(0);
                setTotalResponses(0);
            }
        }

        function filterSurveys() {
            try {
                const now = new Date();
                const filtered = surveys.filter(s => {
                    if (searchQuery && !s.title.toLowerCase().includes(searchQuery.toLowerCase())) return false;
                    if (statusFilters.All) return s.archived !== true;
                    if (statusFilters.draft && s.status === 'draft') return s.archived !== true;
                    if (statusFilters.published && s.status === 'published') return s.archived !== true;
                    if (statusFilters.Upcoming && s.startDate && new Date(s.startDate) > now) return s.archived !== true;
                    if (statusFilters.Active && s.startDate && s.endDate &&
                        new Date(s.startDate) <= now && new Date(s.endDate) >= now) return s.archived !== true;
                    if (statusFilters.Expired && s.endDate && new Date(s.endDate) < now) return s.archived !== true;
                    if (statusFilters.Archived && s.archived === true) return true;
                    return false;
                });
                setFilteredSurveys(filtered);
            } catch (e) {
                console.error('Filter error:', e);
                setFilteredSurveys(surveys);
            }
        }

        React.useEffect(() => {
            filterSurveys();
        }, [statusFilters, searchQuery, surveys]);

        function handleStatusFilterChange(filter) {
            setStatusFilters(prev => ({
                ...Object.fromEntries(Object.keys(prev).map(k => [k, false])),
                [filter]: !prev[filter]
            }));
        }

        function handleSearchChange(e) {
            setSearchQuery(e.target.value);
        }

        function showShare(id, title) {
            setShareLink(`${window.location.origin}/sites/yoursite/FormFill.aspx?surveyId=${id}`);
            setShowShareDialog(true);
            setTimeout(() => {
                if (qrContainerRef.current) {
                    new QRious({
                        element: qrContainerRef.current,
                        value: `${window.location.origin}/sites/yoursite/FormFill.aspx?surveyId=${id}`,
                        size: 200
                    });
                }
            }, 0);
        }

        function downloadQR() {
            const canvas = qrContainerRef.current;
            if (canvas) {
                const link = document.createElement('a');
                link.href = canvas.toDataURL('image/png');
                link.download = 'survey-qr.png';
                link.click();
            }
        }

        function copyURL() {
            navigator.clipboard.writeText(shareLink).then(() => {
                alert('Link copied to clipboard!');
            }).catch(e => {
                console.error('Copy error:', e);
                alert('Failed to copy link.');
            });
        }

        function showDuplicate(id) {
            if (typeof id !== 'number' || id <= 0) {
                console.error('Invalid ID for duplicate:', id);
                alert('Invalid survey ID.');
                return;
            }
            setSelectedId(id);
            setShowDuplicateDialog(true);
        }

        function handleDuplicate() {
            if (selectedId) {
                duplicate(selectedId);
                setShowDuplicateDialog(false);
            }
        }

        function showEditMetadata(survey) {
            setSelectedSurvey({
                ...survey,
                ownerIds: survey.owner.split(', ').map(name => {
                    const user = users.find(u => u.title === name);
                    return user ? user.id : null;
                }).filter(id => id !== null)
            });
            setShowEditMetadataDialog(true);
        }

        function saveMetadata(updatedSurvey) {
            try {
                if (!updatedSurvey.title || updatedSurvey.title.trim() === '') {
                    alert('Title is required.');
                    return;
                }
                if (!updatedSurvey.ownerIds || updatedSurvey.ownerIds.length === 0) {
                    alert('At least one owner is required.');
                    return;
                }
                if (updatedSurvey.startDate && updatedSurvey.endDate &&
                    new Date(updatedSurvey.startDate) > new Date(updatedSurvey.endDate)) {
                    alert('End Date must be after Start Date.');
                    return;
                }
                const context = SP.ClientContext.get_current();
                const surveyList = context.get_web().get_lists().getByTitle('Surveys');
                const item = surveyList.getItemById(updatedSurvey.id);
                context.load(item);
                context.executeQueryAsync(
                    () => {
                        item.set_item('Title', updatedSurvey.title);
                        item.set_item('Status', updatedSurvey.status);
                        item.set_item('StartDate', updatedSurvey.startDate ? new Date(updatedSurvey.startDate) : null);
                        item.set_item('EndDate', updatedSurvey.endDate ? new Date(updatedSurvey.endDate) : null);
                        item.set_item('Archived', updatedSurvey.archived);
                        const ownerFieldValues = updatedSurvey.ownerIds.map(id => {
                            const userFieldValue = new SP.FieldUserValue();
                            userFieldValue.set_lookupId(id);
                            return userFieldValue;
                        });
                        item.set_item('Owner', ownerFieldValues);
                        item.update();
                        const usersToSet = updatedSurvey.ownerIds.map(id => {
                            const user = context.get_web().get_siteUsers().getById(id);
                            context.load(user);
                            return user;
                        });
                        context.executeQueryAsync(
                            () => {
                                setPermissions(item, usersToSet, () => {
                                    console.log('Metadata saved for survey:', updatedSurvey.id); // Governance: Log save
                                    alert('Metadata saved!');
                                    loadSurveysData();
                                    setShowEditMetadataDialog(false);
                                });
                            },
                            onFail('Failed to load users for permissions')
                        );
                    },
                    onFail('Failed to load survey for metadata update')
                );
            } catch (e) {
                console.error('Metadata save error:', e);
                alert('Error saving metadata: ' + e.message);
            }
        }

        return React.createElement('div', { className: 'min-h-screen bg-gray-100' },
            React.createElement('style', null, `
                @media (max-width: 576px) {
                    .modal-responsive {
                        width: 95%;
                        margin: 1rem auto;
                    }
                    .modal-responsive input,
                    .modal-responsive select {
                        font-size: 0.875rem;
                    }
                    .modal-responsive button {
                        font-size: 0.875rem;
                        padding: 0.5rem 1rem;
                    }
                    .card-text {
                        font-size: 0.875rem;
                    }
                    .card-buttons button,
                    .card-buttons a {
                        font-size: 0.75rem;
                        padding: 0.25rem 0.5rem;
                    }
                }
                @media (min-width: 576px) {
                    .modal-responsive {
                        max-width: 500px;
                    }
                }
                @media (min-width: 992px) {
                    .modal-responsive {
                        max-width: 600px;
                    }
                }
                .sidebar {
                    width: 250px;
                }
                .card-buttons {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 0.5rem;
                }
                .card-buttons button,
                .card-buttons a {
                    transition: transform 0.2s, filter 0.2s;
                }
                .card-buttons button:hover,
                .card-buttons a:hover {
                    transform: scale(1.05);
                    filter: brightness(1.1);
                }
                .card-buttons button:focus,
                .card-buttons a:focus {
                    outline: 2px solid #1e40af;
                    outline-offset: 2px;
                    box-shadow: 0 0 0 3px rgba(30, 64, 175, 0.3);
                }
                .multi-select {
                    height: 150px;
                }
                .modal-header {
                    position: relative;
                }
                .modal-close {
                    position: absolute;
                    top: 8px;
                    right: 8px;
                    padding: 0.5rem;
                }
            `),
            React.createElement('nav', { className: 'bg-blue-600 text-white p-4' },
                React.createElement('div', { className: 'container mx-auto flex justify-between items-center' },
                    React.createElement('h1', { className: 'text-2xl font-bold' }, 'Survey Dashboard'),
                    React.createElement('div', null,
                        React.createElement('span', { className: 'mr-4' }, `Welcome, ${username}`),
                        isManager && React.createElement('button', {
                            className: 'bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600 focus:ring-2 focus:ring-offset-2 focus:ring-green-500',
                            onClick: () => window.location.href = '/sites/yoursite/Builder.aspx',
                            'aria-label': 'Create new survey'
                        }, 'Create Survey')
                    )
                )
            ),
            React.createElement('div', { className: 'container mx-auto p-4' },
                React.createElement('div', { className: 'flex flex-col md:flex-row' },
                    React.createElement('div', { className: 'sidebar bg-white shadow-md p-4 md:mr-4 mb-4 md:mb-0' },
                        React.createElement('h2', { className: 'text-lg font-bold mb-4' }, 'Filters'),
                        Object.keys(statusFilters).map(filter => React.createElement('div', { key: filter, className: 'mb-2' },
                            React.createElement('input', {
                                type: 'checkbox',
                                checked: statusFilters[filter],
                                onChange: () => handleStatusFilterChange(filter),
                                className: 'mr-2',
                                'aria-label': `Filter by ${filter}`
                            }),
                            filter
                        )),
                        React.createElement('input', {
                            type: 'text',
                            placeholder: 'Search surveys...',
                            value: searchQuery,
                            onChange: handleSearchChange,
                            className: 'mt-4 w-full p-2 border rounded focus:ring-2 focus:ring-blue-500',
                            'aria-label': 'Search surveys'
                        })
                    ),
                    React.createElement('div', { className: 'flex-1 p-4' },
                        React.createElement('div', { className: 'mb-4' }, // Removed Export CSV button
                            React.createElement('div', null,
                                React.createElement('p', null, `Total Surveys: ${totalSurveys}`),
                                React.createElement('p', null, `Total Responses: ${totalResponses}`)
                            )
                        ),
                        React.createElement('div', { className: 'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4' },
                            filteredSurveys.length === 0 ?
                                React.createElement('p', { className: 'text-gray-500' }, 'No surveys found.') :
                                filteredSurveys.map(survey => React.createElement('div', {
                                    key: survey.id,
                                    className: 'bg-white shadow-md rounded-lg p-4 h-full'
                                },
                                    React.createElement('div', {
                                        className: 'flex flex-col gap-2',
                                        'aria-label': `Survey: ${survey.title}`
                                    },
                                        React.createElement('h5', { className: 'text-lg font-bold' }, survey.title),
                                        React.createElement('p', { className: 'card-text text-gray-700' }, 'Owner(s): ', survey.owner),
                                        React.createElement('p', { className: 'card-text text-gray-700' }, 'Created: ', survey.created),
                                        React.createElement('p', { className: 'card-text text-gray-700' }, 'Status: ', survey.status),
                                        React.createElement('p', { className: 'card-text text-gray-700' }, 'Start Date: ', survey.startDate),
                                        React.createElement('p', { className: 'card-text text-gray-700' }, 'End Date: ', survey.endDate),
                                        React.createElement('p', { className: 'card-text text-gray-700' }, 'Archived: ', survey.archived ? 'Yes' : 'No'),
                                        React.createElement('p', { className: 'card-text text-gray-700' }, 'Responses: ', survey.responses),
                                        React.createElement('div', { className: 'card-buttons flex flex-wrap gap-2' },
                                            React.createElement('a', {
                                                href: `/sites/yoursite/Builder.aspx?surveyId=${survey.id}`,
                                                target: '_blank',
                                                rel: 'noopener noreferrer',
                                                className: 'bg-blue-500 text-white px-2 py-1 rounded text-sm hover:scale-105 focus:ring-2 focus:ring-offset-2 focus:ring-blue-500',
                                                'aria-label': `Edit survey ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-pencil-alt mr-1' }), 'Edit'),
                                            React.createElement('a', {
                                                href: `/sites/yoursite/FormFill.aspx?surveyId=${survey.id}`,
                                                target: '_blank',
                                                rel: 'noopener noreferrer',
                                                className: 'bg-green-500 text-white px-2 py-1 rounded text-sm hover:scale-105 focus:ring-2 focus:ring-offset-2 focus:ring-green-500',
                                                'aria-label': `Run survey ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-play mr-1' }), 'Run'),
                                            React.createElement('a', {
                                                href: `/sites/yoursite/ResponsesView.aspx?surveyId=${survey.id}`,
                                                target: '_blank',
                                                rel: 'noopener noreferrer',
                                                className: 'bg-teal-500 text-white px-2 py-1 rounded text-sm hover:scale-105 focus:ring-2 focus:ring-offset-2 focus:ring-teal-500',
                                                'aria-label': `View responses for ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-chart-bar mr-1' }), 'Responses'),
                                            React.createElement('button', {
                                                className: 'bg-yellow-500 text-white px-2 py-1 rounded text-sm hover:scale-105 focus:ring-2 focus:ring-offset-2 focus:ring-yellow-500',
                                                onClick: () => showShare(survey.id, survey.title),
                                                'aria-label': `Share survey ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-share-alt mr-1' }), 'Share/QR'),
                                            React.createElement('button', {
                                                className: 'bg-gray-500 text-white px-2 py-1 rounded text-sm hover:scale-105 focus:ring-2 focus:ring-offset-2 focus:ring-gray-500',
                                                onClick: () => showEditMetadata(survey),
                                                'aria-label': `Edit metadata for ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-cog mr-1' }), 'Edit Metadata'),
                                            React.createElement('button', {
                                                className: 'bg-gray-800 text-white px-2 py-1 rounded text-sm hover:scale-105 focus:ring-2 focus:ring-offset-2 focus:ring-gray-800',
                                                onClick: () => showDuplicate(survey.id),
                                                'aria-label': `Duplicate survey ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-copy mr-1' }), 'Duplicate')
                                        )
                                    )
                                ))
                        )
                    )
                )
            ),
            showDuplicateDialog && React.createElement('div', { className: 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center' },
                React.createElement('div', { className: 'modal-responsive bg-white rounded-lg shadow-lg mx-auto' },
                    React.createElement('div', { className: 'modal-header flex justify-between items-center p-4 border-b border-gray-200 relative' },
                        React.createElement('h5', { className: 'text-lg font-bold' }, 'Confirm Duplicate'),
                        React.createElement('button', {
                            className: 'modal-close text-gray-600 hover:text-gray-800 focus:ring-2 focus:ring-blue-500',
                            type: 'button',
                            onClick: () => setShowDuplicateDialog(false),
                            'aria-label': 'Close'
                        }, React.createElement('svg', {
                            className: 'w-5 h-5',
                            fill: 'none',
                            stroke: 'currentColor',
                            viewBox: '0 0 24 24'
                        }, React.createElement('path', {
                            strokeLinecap: 'round',
                            strokeLinejoin: 'round',
                            strokeWidth: '2',
                            d: 'M6 18L18 6M6 6l12 12'
                        })))
                    ),
                    React.createElement('div', { className: 'p-4' },
                        React.createElement('p', { className: 'text-gray-700' }, 'Duplicate this survey?')
                    ),
                    React.createElement('div', { className: 'flex justify-end gap-2 p-4 border-t border-gray-200' },
                        React.createElement('button', {
                            className: 'bg-gray-400 text-white px-4 py-2 rounded hover:bg-gray-500 focus:ring-2 focus:ring-offset-2 focus:ring-gray-500',
                            onClick: () => setShowDuplicateDialog(false),
                            'aria-label': 'Cancel'
                        }, 'Cancel'),
                        React.createElement('button', {
                            className: 'bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600 focus:ring-2 focus:ring-offset-2 focus:ring-blue-500',
                            onClick: handleDuplicate,
                            'aria-label': 'Confirm duplicate'
                        }, 'Duplicate')
                    )
                )
            ),
            showShareDialog && React.createElement('div', { className: 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center' },
                React.createElement('div', { className: 'modal-responsive bg-white rounded-lg shadow-lg mx-auto' },
                    React.createElement('div', { className: 'modal-header flex justify-between items-center p-4 border-b border-gray-200 relative' },
                        React.createElement('h5', { className: 'text-lg font-bold' }, 'Share Survey'),
                        React.createElement('button', {
                            className: 'modal-close text-gray-600 hover:text-gray-800 focus:ring-2 focus:ring-blue-500',
                            type: 'button',
                            onClick: () => setShowShareDialog(false),
                            'aria-label': 'Close'
                        }, React.createElement('svg', {
                            className: 'w-5 h-5',
                            fill: 'none',
                            stroke: 'currentColor',
                            viewBox: '0 0 24 24'
                        }, React.createElement('path', {
                            strokeLinecap: 'round',
                            strokeLinejoin: 'round',
                            strokeWidth: '2',
                            d: 'M6 18L18 6M6 6l12 12'
                        })))
                    ),
                    React.createElement('div', { className: 'p-4' },
                        React.createElement('p', { className: 'text-gray-700 mb-2' }, 'Share this link:'),
                        React.createElement('input', {
                            type: 'text',
                            value: shareLink,
                            readOnly: true,
                            className: 'w-full p-2 border rounded mb-2 focus:ring-2 focus:ring-blue-500',
                            'aria-label': 'Survey share link'
                        }),
                        React.createElement('canvas', { ref: qrContainerRef, className: 'mb-2' }),
                        React.createElement('div', { className: 'flex gap-2' },
                            React.createElement('button', {
                                className: 'bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600 focus:ring-2 focus:ring-offset-2 focus:ring-blue-500',
                                onClick: copyURL,
                                'aria-label': 'Copy share link'
                            }, 'Copy Link'),
                            React.createElement('button', {
                                className: 'bg-green-500 text-white px-4 py-2 rounded hover:bg-green-600 focus:ring-2 focus:ring-offset-2 focus:ring-green-500',
                                onClick: downloadQR,
                                'aria-label': 'Download QR code'
                            }, 'Download QR')
                        )
                    )
                )
            ),
            showEditMetadataDialog && React.createElement('div', { className: 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center' },
                React.createElement('div', { className: 'modal-responsive bg-white rounded-lg shadow-lg mx-auto' },
                    React.createElement('div', { className: 'modal-header flex justify-between items-center p-4 border-b border-gray-200 relative' },
                        React.createElement('h5', { className: 'text-lg font-bold' }, 'Edit Metadata'),
                        React.createElement('button', {
                            className: 'modal-close text-gray-600 hover:text-gray-800 focus:ring-2 focus:ring-blue-500',
                            type: 'button',
                            onClick: () => setShowEditMetadataDialog(false),
                            'aria-label': 'Close'
                        }, React.createElement('svg', {
                            className: 'w-5 h-5',
                            fill: 'none',
                            stroke: 'currentColor',
                            viewBox: '0 0 24 24'
                        }, React.createElement('path', {
                            strokeLinecap: 'round',
                            strokeLinejoin: 'round',
                            strokeWidth: '2',
                            d: 'M6 18L18 6M6 6l12 12'
                        })))
                    ),
                    React.createElement('div', { className: 'p-4' },
                        React.createElement('div', { className: 'mb-4' },
                            React.createElement('label', { className: 'block text-gray-700' }, 'Title'),
                            React.createElement('input', {
                                type: 'text',
                                value: selectedSurvey?.title || '',
                                onChange: e => setSelectedSurvey({ ...selectedSurvey, title: e.target.value }),
                                className: 'w-full p-2 border rounded focus:ring-2 focus:ring-blue-500',
                                'aria-label': 'Survey title'
                            })
                        ),
                        React.createElement('div', { className: 'mb-4' },
                            React.createElement('label', { className: 'block text-gray-700' }, 'Owner(s)'),
                            React.createElement('select', {
                                multiple: true,
                                value: selectedSurvey?.ownerIds || [],
                                onChange: e => setSelectedSurvey({
                                    ...selectedSurvey,
                                    ownerIds: Array.from(e.target.selectedOptions).map(opt => parseInt(opt.value))
                                }),
                                className: 'multi-select w-full p-2 border rounded focus:ring-2 focus:ring-blue-500',
                                'aria-label': 'Select survey owners'
                            }, users.map(user => React.createElement('option', {
                                key: user.id,
                                value: user.id
                            }, user.title)))
                        ),
                        React.createElement('div', { className: 'mb-4' },
                            React.createElement('label', { className: 'block text-gray-700' }, 'Status'),
                            React.createElement('select', {
                                value: selectedSurvey?.status || 'draft',
                                onChange: e => setSelectedSurvey({ ...selectedSurvey, status: e.target.value }),
                                className: 'w-full p-2 border rounded focus:ring-2 focus:ring-blue-500',
                                'aria-label': 'Survey status'
                            }, ['draft', 'published'].map(status => React.createElement('option', {
                                key: status,
                                value: status
                            }, status.charAt(0).toUpperCase() + status.slice(1))))
                        ),
                        React.createElement('div', { className: 'mb-4' },
                            React.createElement('label', { className: 'block text-gray-700' }, 'Start Date'),
                            React.createElement('input', {
                                type: 'date',
                                value: selectedSurvey?.startDate ? new Date(selectedSurvey.startDate).toISOString().split('T')[0] : '',
                                onChange: e => setSelectedSurvey({ ...selectedSurvey, startDate: e.target.value }),
                                className: 'w-full p-2 border rounded focus:ring-2 focus:ring-blue-500',
                                'aria-label': 'Survey start date'
                            })
                        ),
                        React.createElement('div', { className: 'mb-4' },
                            React.createElement('label', { className: 'block text-gray-700' }, 'End Date'),
                            React.createElement('input', {
                                type: 'date',
                                value: selectedSurvey?.endDate ? new Date(selectedSurvey.endDate).toISOString().split('T')[0] : '',
                                onChange: e => setSelectedSurvey({ ...selectedSurvey, endDate: e.target.value }),
                                className: 'w-full p-2 border rounded focus:ring-2 focus:ring-blue-500',
                                'aria-label': 'Survey end date'
                            })
                        ),
                        React.createElement('div', { className: 'mb-4' },
                            React.createElement('label', { className: 'block text-gray-700' }, 'Archived'),
                            React.createElement('input', {
                                type: 'checkbox',
                                checked: selectedSurvey?.archived || false,
                                onChange: e => setSelectedSurvey({ ...selectedSurvey, archived: e.target.checked }),
                                className: 'mr-2',
                                'aria-label': 'Archive survey'
                            }, 'Archived')
                        )
                    ),
                    React.createElement('div', { className: 'flex justify-end gap-2 p-4 border-t border-gray-200' },
                        React.createElement('button', {
                            className: 'bg-gray-400 text-white px-4 py-2 rounded hover:bg-gray-500 focus:ring-2 focus:ring-offset-2 focus:ring-gray-500',
                            onClick: () => setShowEditMetadataDialog(false),
                            'aria-label': 'Cancel'
                        }, 'Cancel'),
                        React.createElement('button', {
                            className: 'bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600 focus:ring-2 focus:ring-offset-2 focus:ring-blue-500',
                            onClick: () => saveMetadata(selectedSurvey),
                            'aria-label': 'Save metadata'
                        }, 'Save')
                    )
                )
            )
        );
    };

    ReactDOM.render(React.createElement(App), document.getElementById('app'));
}

function duplicate(id) {
    try {
        if (typeof id !== 'number' || id <= 0) {
            throw new Error('Invalid ID: Expected a positive number');
        }
        console.log('Duplicate ID:', id); // Governance: Log ID
        const context = SP.ClientContext.get_current();
        const web = context.get_web();
        const surveyList = web.get_lists().getByTitle('Surveys');
        context.load(surveyList);
        context.executeQueryAsync(
            () => {
                console.log('Surveys list loaded successfully'); // Governance: Log list access
                const item = surveyList.getItemById(id);
                context.load(item, 'Include(Id, Title, surveyData, Owner, Status, StartDate, EndDate, Archived)');
                context.executeQueryAsync(
                    () => {
                        console.log('Item loaded:', item.get_fieldValues()); // Governance: Log item
                        try {
                            const surveyData = item.get_item('surveyData');
                            console.log('surveyData retrieved:', surveyData); // Governance: Log JSON
                            let json;
                            if (!surveyData) {
                                console.warn('surveyData is empty; using default object');
                                json = { title: 'Survey' };
                            } else {
                                try {
                                    json = JSON.parse(surveyData);
                                } catch (parseError) {
                                    console.error('JSON parsing error:', parseError);
                                    throw new Error('Invalid surveyData format');
                                }
                            }
                            if (!json || typeof json !== 'object') {
                                console.warn('surveyData is invalid; using default object');
                                json = { title: 'Survey' };
                            }
                            json.title = `Copy of ${json.title || 'Survey'}`;
                            saveSurvey(json, null, success => {
                                if (success) {
                                    console.log('Survey duplicated, ID:', id); // Governance: Log success
                                    alert('Survey duplicated!');
                                    location.reload();
                                } else {
                                    console.error('Failed to save duplicated survey');
                                    alert('Error duplicating survey: Failed to save');
                                }
                            });
                        } catch (e) {
                            console.error('Duplicate JSON error:', e);
                            alert('Error duplicating survey: ' + e.message);
                        }
                    },
                    (sender, args) => {
                        console.error('Failed to load survey item:', {
                            id: id,
                            errorMessage: args.get_message(),
                            errorStack: args.get_stackTrace ? args.get_stackTrace() : 'No stack trace'
                        });
                        alert(`Failed to load survey for duplication: ${args.get_message()}`);
                    }
                );
            },
            (sender, args) => {
                console.error('Failed to load Surveys list:', {
                    errorMessage: args.get_message(),
                    errorStack: args.get_stackTrace ? args.get_stackTrace() : 'No stack trace'
                });
                alert(`Failed to access Surveys list: ${args.get_message()}`);
            }
        );
    } catch (e) {
        console.error('Duplicate error:', e);
        alert('Error duplicating survey: ' + e.message);
    }
}

function saveSurvey(json, id, callback) {
    try {
        if (!json || typeof json !== 'object') throw new Error('Invalid JSON');
        console.log('Saving survey, ID:', id || 'New'); // Governance: Log save
        const context = SP.ClientContext.get_current();
        const web = context.get_web();
        const user = web.get_currentUser();
        const surveyList = web.get_lists().getByTitle('Surveys');
        context.load(user);
        context.executeQueryAsync(
            () => {
                let item;
                if (id) {
                    item = surveyList.getItemById(id);
                    item.set_item('surveyData', JSON.stringify(json));
                    item.set_item('Title', json.title || 'Updated Survey');
                } else {
                    const createInfo = new SP.ListItemCreationInformation();
                    item = surveyList.addItem(createInfo);
                    item.set_item('Title', json.title || 'New Survey');
                    item.set_item('surveyData', JSON.stringify(json));
                    const userFieldValue = new SP.FieldUserValue();
                    userFieldValue.set_lookupId(user.get_id());
                    item.set_item('Owner', [userFieldValue]);
                    item.set_item('Status', 'draft');
                    item.set_item('Archived', false);
                }
                item.update();
                context.load(item);
                context.executeQueryAsync(
                    () => {
                        if (!id) {
                            setPermissions(item, [user], () => {
                                console.log('Survey created, ID:', item.get_id()); // Governance: Log creation
                                callback(true);
                            });
                        } else {
                            console.log('Survey updated, ID:', id); // Governance: Log update
                            callback(true);
                        }
                    },
                    (sender, args) => {
                        console.error('Failed to save survey:', {
                            id: id,
                            errorMessage: args.get_message(),
                            errorStack: args.get_stackTrace ? args.get_stackTrace() : 'No stack trace'
                        });
                        alert(`Failed to save survey: ${args.get_message()}`);
                        callback(false);
                    }
                );
            },
            (sender, args) => {
                console.error('Failed to load user:', {
                    errorMessage: args.get_message(),
                    errorStack: args.get_stackTrace ? args.get_stackTrace() : 'No stack trace'
                });
                alert(`Failed to load user: ${args.get_message()}`);
                callback(false);
            }
        );
    } catch (e) {
        console.error('Save error:', e);
        alert('Error saving survey: ' + e.message);
        callback(false);
    }
}

function setPermissions(item, users, callback) {
    try {
        console.log('Setting permissions for survey:', item.get_id()); // Governance: Log permissions
        const context = SP.ClientContext.get_current();
        const web = context.get_web();
        users.forEach(user => context.load(user));
        item.breakRoleInheritance(true, false);
        const roleAssignments = item.get_roleAssignments();
        context.load(roleAssignments);
        context.executeQueryAsync(
            () => {
                const assignments = roleAssignments.get_data();
                const toRemove = [];
                assignments.forEach(assignment => {
                    const principal = assignment.get_member();
                    if (principal.get_principalType() === SP.PrincipalType.user) {
                        toRemove.push(assignment);
                    }
                });
                toRemove.forEach(assignment => {
                    assignment.deleteObject();
                });
                const roleDefs = web.get_roleDefinitions();
                const contribute = roleDefs.getByName('Contribute');
                users.forEach(user => {
                    const newAssignment = new SP.RoleAssignment(user);
                    newAssignment.addRoleDefinition(contribute);
                    roleAssignments.add(newAssignment);
                });
                item.update();
                context.executeQueryAsync(
                    () => {
                        console.log('Permissions set for users:', users.map(u => u.get_title())); // Governance: Log success
                        callback();
                    },
                    (sender, args) => {
                        console.error('Failed to set permissions:', {
                            errorMessage: args.get_message(),
                            errorStack: args.get_stackTrace ? args.get_stackTrace() : 'No stack trace'
                        });
                        alert(`Failed to set permissions: ${args.get_message()}`);
                        callback();
                    }
                );
            },
            (sender, args) => {
                console.error('Failed to load role assignments:', {
                    errorMessage: args.get_message(),
                    errorStack: args.get_stackTrace ? args.get_stackTrace() : 'No stack trace'
                });
                alert(`Failed to load role assignments: ${args.get_message()}`);
                callback();
            }
        );
    } catch (e) {
        console.error('Permissions error:', e);
        alert('Error setting permissions: ' + e.message);
        callback();
    }
}

function onFail(message) {
    return (sender, args) => {
        console.error(message, args.get_message());
        alert(`${message}: ${args.get_message()}`);
    };
}