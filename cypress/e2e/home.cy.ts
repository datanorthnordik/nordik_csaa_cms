describe("Home page", () => {
  it("loads successfully", () => {
    cy.visit("/");
    // Check if page loads without errors
    cy.get("body").should("exist");
  });
});
