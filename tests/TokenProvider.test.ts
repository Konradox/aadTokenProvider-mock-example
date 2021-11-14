///<reference types="jest" />
import { assert } from "chai";
import { TokenProvider } from "../src/TokenProvider";
import { SPWebPartContextMock } from "spfx-ut-library";
import { Log } from "@microsoft/sp-core-library";

jest.mock("@microsoft/sp-core-library", () => ({
	Log: {
		error: jest.fn(),
	},
}));

describe("TokenProvider", () => {
	let mockContext = new SPWebPartContextMock();

	afterEach(() => {
		mockContext.aadTokenProviderFactory.aadTokenProviderMock.clearMocks();
		(Log.error as jest.Mock).mockClear();
	});

	test("Should return graph token", async () => {
		const TOKEN = "token";
		mockContext.aadTokenProviderFactory.aadTokenProviderMock.registerToken("https://graph.microsoft.com", TOKEN);
		const tokenProvider = new TokenProvider(mockContext as any);
		const token = await tokenProvider.getToken("https://graph.microsoft.com");
		assert.equal(token, TOKEN);
	});

	test("Should throw error with errorCode = invalid_resource when resource is unavailable", async () => {
		const tokenProvider = new TokenProvider(mockContext as any);
		(Log.error as jest.Mock).mockImplementation((source, error) => {
			assert.equal(source, "TokenProvider");
			assert.equal(error.errorCode, "invalid_resource");
		});
		try {
			await tokenProvider.getToken("https://fabricated.service.microsoft.com");
			assert.fail("Should throw error");
		} catch (error) {
			assert.equal(error.errorCode, "invalid_resource");
		}
	});

	test("Should throw custom error", async () => {
		mockContext.aadTokenProviderFactory.aadTokenProviderMock.registerError("https://fabricated.service.microsoft.com", {
			errorCode: "custom_error",
			errorMessage: "Custom error message",
			message: "Custom error message",
			name: "Custom error name",
			stack: "Custom error stack",
		});
		const tokenProvider = new TokenProvider(mockContext as any);
		(Log.error as jest.Mock).mockImplementation((source, error) => {
			assert.equal(source, "TokenProvider");
			assert.equal(error.errorCode, "custom_error");
			assert.equal(error.errorMessage, "Custom error message");
			assert.equal(error.message, "Custom error message");
			assert.equal(error.name, "Custom error name");
			assert.equal(error.stack, "Custom error stack");
		});
		try {
			await tokenProvider.getToken("https://fabricated.service.microsoft.com");
			assert.fail("Should throw error");
		} catch (error) {
			assert.equal(error.errorCode, "custom_error");
		}
	});
});
