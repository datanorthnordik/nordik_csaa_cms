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
  it('requires audio and creates an item with an uploaded recording', () => {
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
            description: 'A community oral history.',
            recording_url: '/api/recordings/7/items/11/content',
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
      'A community oral history.',
    )
    cy.get('button').contains(/^Create Collection$/).click()
    cy.contains('Each item needs a recording file.').should('be.visible')

    cy.get('input[type="file"]').selectFile(
      {
        contents: Cypress.Buffer.from('ID3 fake audio data'),
        fileName: 'september-gathering.mp3',
        mimeType: 'audio/mpeg',
      },
      { force: true },
    )
    cy.contains('Selected recording: september-gathering.mp3').should('be.visible')
    cy.get('button').contains(/^Create Collection$/).click()

    cy.wait('@createRecordingCollection').then(({ request }) => {
      expect(request.headers['content-type']).to.contain('multipart/form-data')
    })
    cy.wait('@getRecordingCollection')
    cy.location('pathname').should('eq', '/recordings/7')
    cy.get('input[value="September gathering"]').should('be.visible')
    cy.get('audio[controls]').should('have.attr', 'src').and('contain', '/api/recordings/7/items/11/content')
    cy.contains('Drop a replacement recording here or browse').should('be.visible')
  })
})
