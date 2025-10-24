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
                console.error('SP.js failed to load.');
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
    if (!SP.ClientContext) {
        console.error('SP.ClientContext not available');
        alert('SharePoint context failed to load.');
        return;
    }
    if (!React || !ReactDOM) {
        console.error('React or ReactDOM not loaded');
        alert('React scripts failed to load.');
        return;
    }
    if (!window.QRious) {
        console.error('QRious not loaded');
        alert('QRious script failed to load.');
        return;
    }

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
        const [selectedId, setSelectedId] = React.useState(null);
        const [showDuplicateDialog, setShowDuplicateDialog] = React.useState(false);
        const [showShareDialog, setShowShareDialog] = React.useState(false);
        const [showEditMetadataDialog, setShowEditMetadataDialog] = React.useState(false);
        const [selectedSurvey, setSelectedSurvey] = React.useState(null);
        const [users, setUsers] = React.useState([]);
        const [shareLink, setShareLink] = React.useState('');
        const [sidebarOpen, setSidebarOpen] = React.useState(true);
        const qrContainerRef = React.useRef(null);

        React.useEffect(() => {
            loadUserInfo();
            loadUserRole();
            loadSiteUsers();
            loadSurveysData();
            fetchCounts();
            // Handle sidebar responsiveness
            const handleResize = () => setSidebarOpen(window.innerWidth >= 992);
            window.addEventListener('resize', handleResize);
            handleResize();
            return () => window.removeEventListener('resize', handleResize);
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
                const user = context.get_web().get_currentUser();
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
                const users = context.get_web().get_siteUsers();
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
                        context.load(items, 'Include(Id, Title, Owner, Status, StartDate, EndDate, Archived, surveyData, Created)');
                        context.executeQueryAsync(
                            () => {
                                console.log('Surveys loaded successfully');
                                const surveyData = [];
                                const enumerator = items.getEnumerator();
                                while (enumerator.moveNext()) {
                                    const item = enumerator.get_current();
                                    const owners = item.get_item('Owner') || [];
                                    const ownerNames = Array.isArray(owners)
                                        ? owners.map(o => o.get_lookupValue()).join(', ')
                                        : 'Unknown';
                                    surveyData.push({
                                        id: item.get_id(),
                                        title: item.get_item('Title') || 'Untitled',
                                        owners: ownerNames.split(', ').map(name => name.trim()).filter(name => name !== 'Unknown'),
                                        ownerIds: Array.isArray(owners) ? owners.map(o => o.get_lookupId()) : [],
                                        created: item.get_item('Created') ? new Date(item.get_item('Created')).toLocaleDateString() : 'Unknown',
                                        status: item.get_item('Status') || 'draft',
                                        startDate: item.get_item('StartDate') ? new Date(item.get_item('StartDate')).toLocaleDateString() : 'No Start Date',
                                        endDate: item.get_item('EndDate') ? new Date(item.get_item('EndDate')).toLocaleDateString() : 'No End Date',
                                        archived: item.get_item('Archived') || false,
                                        responses: 0
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
                        console.log('Responses counted successfully');
                        const responseCounts = {};
                        const enumerator = items.getEnumerator();
                        while (enumerator.moveNext()) {
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
            const url = `https://yourtenant.sharepoint.com/sites/yoursite/FormFill.aspx?surveyId=${id}`;
            setShareLink(url);
            setSelectedSurvey({ id, title });
            setShowShareDialog(true);
            setTimeout(() => {
                if (qrContainerRef.current) {
                    new QRious({
                        element: qrContainerRef.current,
                        value: url,
                        size: 200
                    });
                }
            }, 0);
        }

        function downloadQR() {
            const canvas = qrContainerRef.current;
            if (canvas && selectedSurvey) {
                const link = document.createElement('a');
                link.href = canvas.toDataURL('image/png');
                link.download = `${selectedSurvey.title.replace(/[^a-z0-9]/gi, '_')}-qr.png`;
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
                startDate: survey.startDate && survey.startDate !== 'No Start Date' ? new Date(survey.startDate).toISOString().split('T')[0] : '',
                endDate: survey.endDate && survey.endDate !== 'No End Date' ? new Date(survey.endDate).toISOString().split('T')[0] : ''
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
                                    console.log('Metadata saved for survey:', updatedSurvey.id);
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
                alert('Error saving metadata.');
            }
        }

        return React.createElement('div', { className: 'min-vh-100 bg-light' },
            React.createElement('nav', { className: 'navbar navbar-dark bg-primary' },
                React.createElement('div', { className: 'container-fluid' },
                    React.createElement('div', { className: 'd-flex align-items-center' },
                        React.createElement('button', {
                            className: 'navbar-toggler me-2 d-lg-none',
                            type: 'button',
                            onClick: () => setSidebarOpen(!sidebarOpen),
                            'aria-label': 'Toggle sidebar'
                        }, React.createElement('span', { className: 'navbar-toggler-icon' })),
                        React.createElement('span', { className: 'navbar-brand' }, 'Survey Dashboard')
                    ),
                    React.createElement('div', null,
                        React.createElement('span', { className: 'text-white me-3' }, `Welcome, ${username}`),
                        isManager && React.createElement('a', {
                            href: '/sites/yoursite/Builder.aspx',
                            className: 'btn btn-success',
                            'aria-label': 'Create new survey'
                        }, 'Create Survey')
                    )
                )
            ),
            React.createElement('div', { className: 'container-fluid' },
                React.createElement('div', { className: 'row' },
                    React.createElement('div', {
                        className: `col-lg-3 bg-white border-end ${sidebarOpen ? '' : 'd-none d-lg-block'}`
                    },
                        React.createElement('div', { className: 'p-3' },
                            React.createElement('h5', null, 'Filters'),
                            Object.keys(statusFilters).map(filter => React.createElement('div', {
                                key: filter,
                                className: 'form-check'
                            },
                                React.createElement('input', {
                                    type: 'checkbox',
                                    className: 'form-check-input',
                                    checked: statusFilters[filter],
                                    onChange: () => handleStatusFilterChange(filter),
                                    id: `filter-${filter}`,
                                    'aria-label': `Filter by ${filter}`
                                }),
                                React.createElement('label', {
                                    className: 'form-check-label',
                                    htmlFor: `filter-${filter}`
                                }, filter)
                            )),
                            React.createElement('input', {
                                type: 'text',
                                placeholder: 'Search surveys...',
                                value: searchQuery,
                                onChange: handleSearchChange,
                                className: 'form-control mt-3',
                                'aria-label': 'Search surveys'
                            })
                        )
                    ),
                    React.createElement('div', { className: 'col-lg-9 p-3' },
                        React.createElement('div', { className: 'mb-3' },
                            React.createElement('p', null, `Total Surveys: ${totalSurveys}`),
                            React.createElement('p', null, `Total Responses: ${totalResponses}`)
                        ),
                        React.createElement('div', { className: 'row row-cols-1 row-cols-md-2 row-cols-lg-3 g-3' },
                            filteredSurveys.length === 0 ?
                                React.createElement('p', { className: 'text-muted' }, 'No surveys found.') :
                                filteredSurveys.map(survey => React.createElement('div', {
                                    key: survey.id,
                                    className: 'col'
                                },
                                    React.createElement('div', {
                                        className: 'card h-100',
                                        'aria-label': `Survey: ${survey.title}`
                                    },
                                        React.createElement('div', { className: 'card-body' },
                                            React.createElement('h5', { className: 'card-title' }, survey.title),
                                            React.createElement('p', { className: 'card-text' },
                                                React.createElement('strong', null, 'Owner(s): '),
                                                survey.owners.length > 0 ? survey.owners.map((owner, idx) => 
                                                    React.createElement('span', { key: idx, className: 'd-block' }, owner)
                                                ) : 'Unknown'
                                            ),
                                            React.createElement('p', { className: 'card-text' }, `Created: ${survey.created}`),
                                            React.createElement('p', { className: 'card-text' }, `Status: ${survey.status}`),
                                            React.createElement('p', { className: 'card-text' }, `Start Date: ${survey.startDate}`),
                                            React.createElement('p', { className: 'card-text' }, `End Date: ${survey.endDate}`),
                                            React.createElement('p', { className: 'card-text' }, `Archived: ${survey.archived ? 'Yes' : 'No'}`),
                                            React.createElement('p', { className: 'card-text' }, `Responses: ${survey.responses}`)
                                        ),
                                        React.createElement('div', { className: 'card-footer d-flex flex-wrap gap-2' },
                                            React.createElement('a', {
                                                href: `/sites/yoursite/Builder.aspx?surveyId=${survey.id}`,
                                                target: '_blank',
                                                rel: 'noopener noreferrer',
                                                className: 'btn btn-primary btn-sm',
                                                'aria-label': `Edit survey ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-pencil-alt me-1' }), 'Edit'),
                                            React.createElement('a', {
                                                href: `/sites/yoursite/FormFill.aspx?surveyId=${survey.id}`,
                                                target: '_blank',
                                                rel: 'noopener noreferrer',
                                                className: 'btn btn-success btn-sm',
                                                'aria-label': `Run survey ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-play me-1' }), 'Run'),
                                            React.createElement('a', {
                                                href: `/sites/yoursite/ResponsesView.aspx?surveyId=${survey.id}`,
                                                target: '_blank',
                                                rel: 'noopener noreferrer',
                                                className: 'btn btn-info btn-sm',
                                                'aria-label': `View responses for ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-chart-bar me-1' }), 'Report'),
                                            React.createElement('button', {
                                                className: 'btn btn-warning btn-sm',
                                                onClick: () => showShare(survey.id, survey.title),
                                                'aria-label': `Share survey ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-share-alt me-1' }), 'Share'),
                                            React.createElement('button', {
                                                className: 'btn btn-secondary btn-sm',
                                                onClick: () => showEditMetadata(survey),
                                                'aria-label': `Edit metadata for ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-cog me-1' }), 'Edit Metadata'),
                                            React.createElement('button', {
                                                className: 'btn btn-dark btn-sm',
                                                onClick: () => showDuplicate(survey.id),
                                                'aria-label': `Duplicate survey ${survey.title}`
                                            }, React.createElement('i', { className: 'fas fa-copy me-1' }), 'Duplicate')
                                        )
                                    )
                                ))
                        )
                    )
                )
            ),
            showDuplicateDialog && React.createElement('div', {
                className: 'modal fade show d-block',
                tabIndex: '-1',
                'aria-labelledby': 'duplicateModalLabel',
                'aria-hidden': !showDuplicateDialog
            },
                React.createElement('div', { className: 'modal-dialog modal-dialog-centered' },
                    React.createElement('div', { className: 'modal-content' },
                        React.createElement('div', { className: 'modal-header' },
                            React.createElement('h5', { className: 'modal-title', id: 'duplicateModalLabel' }, 'Confirm Duplicate'),
                            React.createElement('button', {
                                type: 'button',
                                className: 'btn-close',
                                onClick: () => setShowDuplicateDialog(false),
                                'aria-label': 'Close'
                            })
                        ),
                        React.createElement('div', { className: 'modal-body' },
                            React.createElement('p', null, 'Duplicate this survey?')
                        ),
                        React.createElement('div', { className: 'modal-footer' },
                            React.createElement('button', {
                                className: 'btn btn-secondary',
                                onClick: () => setShowDuplicateDialog(false),
                                'aria-label': 'Cancel'
                            }, 'Cancel'),
                            React.createElement('button', {
                                className: 'btn btn-primary',
                                onClick: handleDuplicate,
                                'aria-label': 'Confirm duplicate'
                            }, 'Duplicate')
                        )
                    )
                )
            ),
            showShareDialog && React.createElement('div', {
                className: 'modal fade show d-block',
                tabIndex: '-1',
                'aria-labelledby': 'shareModalLabel',
                'aria-hidden': !showShareDialog
            },
                React.createElement('div', { className: 'modal-dialog modal-dialog-centered' },
                    React.createElement('div', { className: 'modal-content' },
                        React.createElement('div', { className: 'modal-header' },
                            React.createElement('h5', { className: 'modal-title', id: 'shareModalLabel' }, 'Share Survey'),
                            React.createElement('button', {
                                type: 'button',
                                className: 'btn-close',
                                onClick: () => setShowShareDialog(false),
                                'aria-label': 'Close'
                            })
                        ),
                        React.createElement('div', { className: 'modal-body' },
                            React.createElement('p', null, 'Share this link:'),
                            React.createElement('input', {
                                type: 'text',
                                value: shareLink,
                                readOnly: true,
                                className: 'form-control mb-2',
                                'aria-label': 'Survey share link'
                            }),
                            React.createElement('canvas', { ref: qrContainerRef, className: 'mb-2 d-block mx-auto' }),
                            React.createElement('div', { className: 'd-flex gap-2' },
                                React.createElement('button', {
                                    className: 'btn btn-primary',
                                    onClick: copyURL,
                                    'aria-label': 'Copy share link'
                                }, 'Copy URL'),
                                React.createElement('button', {
                                    className: 'btn btn-success',
                                    onClick: downloadQR,
                                    'aria-label': 'Download QR code'
                                }, 'Download QR')
                            )
                        )
                    )
                )
            ),
            showEditMetadataDialog && React.createElement('div', {
                className: 'modal fade show d-block',
                tabIndex: '-1',
                'aria-labelledby': 'editMetadataModalLabel',
                'aria-hidden': !showEditMetadataDialog
            },
                React.createElement('div', { className: 'modal-dialog modal-dialog-centered' },
                    React.createElement('div', { className: 'modal-content' },
                        React.createElement('div', { className: 'modal-header' },
                            React.createElement('h5', { className: 'modal-title', id: 'editMetadataModalLabel' }, 'Edit Metadata'),
                            React.createElement('button', {
                                type: 'button',
                                className: 'btn-close',
                                onClick: () => setShowEditMetadataDialog(false),
                                'aria-label': 'Close'
                            })
                        ),
                        React.createElement('div', { className: 'modal-body' },
                            React.createElement('div', { className: 'mb-3' },
                                React.createElement('label', { className: 'form-label' }, 'Title'),
                                React.createElement('input', {
                                    type: 'text',
                                    value: selectedSurvey?.title || '',
                                    onChange: e => setSelectedSurvey({ ...selectedSurvey, title: e.target.value }),
                                    className: 'form-control',
                                    'aria-label': 'Survey title'
                                })
                            ),
                            React.createElement('div', { className: 'mb-3' },
                                React.createElement('label', { className: 'form-label' }, 'Owner(s)'),
                                React.createElement('select', {
                                    multiple: true,
                                    value: selectedSurvey?.ownerIds || [],
                                    onChange: e => setSelectedSurvey({
                                        ...selectedSurvey,
                                        ownerIds: Array.from(e.target.selectedOptions).map(opt => parseInt(opt.value))
                                    }),
                                    className: 'form-select',
                                    style: { height: '150px' },
                                    'aria-label': 'Select survey owners'
                                }, users.map(user => React.createElement('option', {
                                    key: user.id,
                                    value: user.id
                                }, user.title)))
                            ),
                            React.createElement('div', { className: 'mb-3' },
                                React.createElement('label', { className: 'form-label' }, 'Status'),
                                React.createElement('select', {
                                    value: selectedSurvey?.status || 'draft',
                                    onChange: e => setSelectedSurvey({ ...selectedSurvey, status: e.target.value }),
                                    className: 'form-select',
                                    'aria-label': 'Survey status'
                                }, ['draft', 'published'].map(status => React.createElement('option', {
                                    key: status,
                                    value: status
                                }, status.charAt(0).toUpperCase() + status.slice(1))))
                            ),
                            React.createElement('div', { className: 'mb-3' },
                                React.createElement('label', { className: 'form-label' }, 'Start Date'),
                                React.createElement('input', {
                                    type: 'date',
                                    value: selectedSurvey?.startDate || '',
                                    onChange: e => setSelectedSurvey({ ...selectedSurvey, startDate: e.target.value }),
                                    className: 'form-control',
                                    'aria-label': 'Survey start date'
                                })
                            ),
                            React.createElement('div', { className: 'mb-3' },
                                React.createElement('label', { className: 'form-label' }, 'End Date'),
                                React.createElement('input', {
                                    type: 'date',
                                    value: selectedSurvey?.endDate || '',
                                    onChange: e => setSelectedSurvey({ ...selectedSurvey, endDate: e.target.value }),
                                    className: 'form-control',
                                    'aria-label': 'Survey end date'
                                })
                            ),
                            React.createElement('div', { className: 'form-check mb-3' },
                                React.createElement('input', {
                                    type: 'checkbox',
                                    checked: selectedSurvey?.archived || false,
                                    onChange: e => setSelectedSurvey({ ...selectedSurvey, archived: e.target.checked }),
                                    className: 'form-check-input',
                                    id: 'archived-checkbox',
                                    'aria-label': 'Archive survey'
                                }),
                                React.createElement('label', { className: 'form-check-label', htmlFor: 'archived-checkbox' }, 'Archived')
                            )
                        ),
                        React.createElement('div', { className: 'modal-footer' },
                            React.createElement('button', {
                                className: 'btn btn-secondary',
                                onClick: () => setShowEditMetadataDialog(false),
                                'aria-label': 'Cancel'
                            }, 'Cancel'),
                            React.createElement('button', {
                                className: 'btn btn-primary',
                                onClick: () => saveMetadata(selectedSurvey),
                                'aria-label': 'Save metadata'
                            }, 'Save')
                        )
                    )
                )
            )
        );
    };

    ReactDOM.render(React.createElement(App), document.getElementById('app'));
}

function duplicate(id) {
    try {
        if (typeof id !== 'number' || id <= 0) throw new Error('Invalid ID');
        console.log('Duplicate ID:', id);
        const context = SP.ClientContext.get_current();
        const surveyList = context.get_web().get_lists().getByTitle('Surveys');
        const item = surveyList.getItemById(id);
        context.load(item, 'Include(Id, Title, surveyData, Owner, Status, StartDate, EndDate, Archived)');
        context.executeQueryAsync(
            () => {
                console.log('Item loaded:', item.get_fieldValues());
                const surveyData = item.get_item('surveyData');
                let json = surveyData ? JSON.parse(surveyData) : { title: 'Survey' };
                json.title = `Copy of ${json.title || 'Survey'}`;
                saveSurvey(json, null, success => {
                    if (success) {
                        console.log('Survey duplicated, ID:', id);
                        alert('Survey duplicated!');
                        location.reload();
                    } else {
                        console.error('Failed to save duplicated survey');
                        alert('Error duplicating survey.');
                    }
                });
            },
            onFail('Failed to load survey for duplication')
        );
    } catch (e) {
        console.error('Duplicate error:', e);
        alert('Error duplicating survey.');
    }
}

function saveSurvey(json, id, callback) {
    try {
        if (!json || typeof json !== 'object') throw new Error('Invalid JSON');
        console.log('Saving survey, ID:', id || 'New');
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
                                console.log('Survey created, ID:', item.get_id());
                                callback(true);
                            });
                        } else {
                            console.log('Survey updated, ID:', id);
                            callback(true);
                        }
                    },
                    onFail('Failed to save survey')
                );
            },
            onFail('Failed to load user')
        );
    } catch (e) {
        console.error('Save error:', e);
        alert('Error saving survey.');
        callback(false);
    }
}

function setPermissions(item, users, callback) {
    try {
        console.log('Setting permissions for survey:', item.get_id());
        const context = SP.ClientContext.get_current();
        const web = context.get_web();
        users.forEach(user => context.load(user));
        item.breakRoleInheritance(true, false);
        const roleAssignments = item.get_roleAssignments();
        context.load(roleAssignments);
        context.executeQueryAsync(
            () => {
                const assignments = roleAssignments.get_data();
                const toRemove = assignments.filter(a => a.get_member().get_principalType() === SP.PrincipalType.user);
                toRemove.forEach(a => a.deleteObject());
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
                        console.log('Permissions set for users:', users.map(u => u.get_title()));
                        callback();
                    },
                    onFail('Failed to set permissions')
                );
            },
            onFail('Failed to load role assignments')
        );
    } catch (e) {
        console.error('Permissions error:', e);
        alert('Error setting permissions.');
        callback();
    }
}

function onFail(message) {
    return (sender, args) => {
        console.error(message, args.get_message());
        alert(`${message}: ${args.get_message()}`);
    };
}