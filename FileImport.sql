
CREATE PROCEDURE FileImport
	@UserID uniqueidentifier,
	@Name nvarchar(250),
	@Address nvarchar(250)
AS
BEGIN

	SET NOCOUNT ON

	INSERT INTO [User](
		UserID,
		Name,
		[Address])
		VALUES (
		@UserID,
		@Name,
		@Address)
		SELECT SCOPE_IDENTITY() As InsertedID
  
END
GO
