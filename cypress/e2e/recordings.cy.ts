const authState = {
  session: {
    accessToken: 'test-access-token',
    refreshToken: 'test-refresh-token',
    firstname: 'Test',
    lastname: 'Admin',
    id: 1,
    email: 'admin@example.com',
    role: 'admin',
  },
  rememberMe: true,
}

describe('Recordings administration', () => {
  it('creates a collection containing a title and description without a recording', () => {
    cy.intercept('GET', '**/api/recordings', {
      statusCode: 200,
      body: { items: [] },
    }).as('listRecordings')
    cy.intercept('POST', '**/api/recordings', {
      statusCode: 201,
      body: {
        message: 'Recording collection created successfully',
        recording: { id: 7, name: 'Community Gatherings' },
      },
    }).as('createRecordingCollection')
    cy.intercept('GET', '**/api/recordings/7', {
      statusCode: 200,
      body: {
        id: 7,
        name: 'Community Gatherings',
        item_count: 1,
        items: [
          {
            id: 11,
            recording_collection_id: 7,
            title: 'September gathering',
            description: 'The recording will be added later.',
            recording_url: '',
            sort_order: 0,
            created_at: '2026-09-20T10:00:00Z',
            updated_at: '2026-09-20T10:00:00Z',
          },
        ],
        created_at: '2026-09-20T10:00:00Z',
        updated_at: '2026-09-20T10:00:00Z',
      },
    }).as('getRecordingCollection')

    cy.visit('/recordings', {
      onBeforeLoad(window) {
        window.localStorage.setItem('nordik_csaa_auth', JSON.stringify(authState))
      },
    })
    cy.wait('@listRecordings')

    cy.get('h1').contains('Recordings').should('be.visible')
    cy.get('button').contains('Create New Recording Collection').click()
    cy.location('pathname').should('eq', '/recordings/new')

    cy.get('input[placeholder="Enter a collection name..."]').type('Community Gatherings')
    cy.get('button').contains('Add Another Item').click()
    cy.get('input[placeholder="Enter an item title..."]').type('September gathering')
    cy.get('textarea[placeholder="Optional supporting description..."]').type(
      'The recording will be added later.',
    )
    cy.get('button').contains(/^Create Collection$/).click()

    cy.wait('@createRecordingCollection').then(({ request }) => {
      expect(request.body).to.deep.equal({
        name: 'Community Gatherings',
        items: [
          {
            title: 'September gathering',
            description: 'The recording will be added later.',
          },
        ],
      })
    })
    cy.wait('@getRecordingCollection')
    cy.location('pathname').should('eq', '/recordings/7')
    cy.get('input[value="September gathering"]').should('be.visible')
  })
})
